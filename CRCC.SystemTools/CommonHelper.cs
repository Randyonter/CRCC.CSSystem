/****************************************************************************************************************************************************************
 *
 * 版权所有: 2023  All Rights Reserved.
 * 文 件 名:  CommonHelper
 * 描    述:
 * 创 建 者：  Ltf
 * 创建时间：  2023/3/29 15:51:25
 * 历史更新记录:
 * =============================================================================================
*修改标记
*修改时间：2023/3/29 15:51:25
*修改人： Ltf
*版本号： V1.0.0.0
*描述：
 * =============================================================================================

****************************************************************************************************************************************************************/

using CRCC.SystemTools.Log;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Windows.Media.Imaging;

namespace CRCC.SystemTools
{
    public static class CommonHelper
    {
        [Description("获取枚举值")]
        public static List<string> GetEnumValues<T>()
        {
            return Enum.GetNames(typeof(T)).ToList();
        }

        [Description("获取枚举项")]
        public static List<T> GetEnumItems<T>()
        {
            return GetEnumValues<T>().Select(s => s.ToEnum<T>()).ToList();
        }

        [Description("根据描述获取枚举值")]
        public static T GetEnumByDescription<T>(string description, T defaultValue = default(T))
        {
            if (typeof(T).IsEnum == false)
                return defaultValue;
            var res = GetEnumItems<T>().FirstOrDefault(s => (s as Enum)?.GetDescription() == description);
            return res;
        }

        [Description("根据两组对象中的数量获取相似度（返回相似度0-1）(srcCount:第一组中对象总数量;desCount:第二组中对象总数量;sameCount:两组中相同对象的数量;" +
            "sameK:相同数权重（该值越大，相同数越多时相似度越高）;srcMoreK:第一组不同数权重（该值越大，当第一组中存在多余内容时，相似度越低）;" +
            "desMoreK:第二组不同数权重（该值越大，当第二组中存在多余内容时，相似度越低）)")]
        public static double GetSimilarityByCount(double srcCount, double desCount, double sameCount, int sameK = 5, int srcMoreK = 1, int desMoreK = 1)
        {
            if (srcCount <= 0 || desCount <= 0 || sameCount <= 0 || sameCount > srcCount || sameCount > desCount || sameK <= 0)
                return 0;
            var more1 = srcCount - sameCount;
            var more2 = desCount - sameCount;
            return sameCount * sameK / (double)(sameCount * sameK + more1 * srcMoreK + more2 * desMoreK);
        }

        [Description("创建一个Item表，列为：Key和Value")]
        public static DataTable CreateItemTable(bool addEmpty = false)
        {
            var tb = new DataTable();
            tb.Columns.Add(new DataColumn("Key"));
            tb.Columns.Add(new DataColumn("Value"));
            if (addEmpty)
            {
                var datarow = tb.NewRow();
                datarow["Key"] = datarow["Value"] = "";
                tb.Rows.Add(datarow);
            }
            return tb;
        }

        [Description("格式化公式表达式（将用户填写的表达式格式化为标准表达式，并替换参数）")]
        public static string FormatExpression(string exp, Dictionary<string, double> pars = null)
        {
            if (string.IsNullOrEmpty(exp))
                return exp;
            var atts = pars?.Select(s => "var " + s.Key + "=" + s.Value).ToList();
            var parstr = string.Join(";", atts ?? new List<string>());
            var res = exp.Replace("Pow", "Math.pow").Replace("Sqr", "Math.sqrt").Replace("Sin", "Math.sin").Replace("Cos", "Math.cos");
            res = res.Replace("Tan", "Math.tan").Replace("ASin", "Math.asin").Replace("ACos", "Math.acos").Replace("ATan", "Math.atan").Replace("Π", "Math.PI");
            return parstr + ";" + res;
        }

        [Description("从给定的数据源中任取指定数量数据的方案，输出所有取法")]
        public static List<List<T>> Combination<T>(List<T> source, int count)
        {
            if (source == null || source.Any() == false || count <= 0 || count > source.Count)
                return new List<List<T>>();
            return combination<T>(source, 0, count, new List<T>());
        }

        private static List<List<T>> combination<T>(List<T> source, int currentIndex, int remain, List<T> currentSelect)
        {
            if (remain == 0)
                return new List<List<T>>() { currentSelect };
            var result = new List<List<T>>();
            if (source.Count - currentIndex < remain)
                return result;
            var newselect = currentSelect.ToList();
            newselect.Add(source[currentIndex]);
            var result1 = combination(source, currentIndex + 1, remain - 1, newselect);
            var result2 = combination(source, currentIndex + 1, remain, currentSelect.ToList());
            result.AddRange(result1);
            result.AddRange(result2);
            return result;
        }

        [Description("将系统中空dwg图纸复制到指定目录下")]
        public static bool CopyEmptyDwgTemplate(string savepath)
        {
            try
            {
                var srcpath = Path.Combine(DataLibrary.TemplateDirectory, "EmptyDwgTemplate.dwg");
                if (File.Exists(srcpath) == false || string.IsNullOrEmpty(savepath))
                    return false;
                File.Copy(srcpath, savepath, true);
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("复制空dwg文件异常！", "复制空dwg文件异常:" + ex);
                return false;
            }
        }

        #region 获取系统图标和路径

        [Description("根据图片名称或路径获取系统中图片")]
        public static Image GetImage(string imageName)
        {
            string path = File.Exists(imageName) ? imageName : Directory.GetFiles(DataLibrary.ImagesDirectory, imageName + ".*").FirstOrDefault();
            if (File.Exists(path) == false)
                return null;
            return ReadImageFile(path);
        }

        [Description("根据图片名称或路径读取系统中图片（不占用本地图片文件）")]
        public static Image ReadImageFile(string imagePath)
        {
            try
            {
                imagePath = imagePath ?? "";
                if (imagePath.Contains("\\") && File.Exists(imagePath) == false)
                    return null;
                imagePath = File.Exists(imagePath) ? imagePath : Directory.GetFiles(DataLibrary.ImagesDirectory, imagePath + ".*").FirstOrDefault();
                if (File.Exists(imagePath) == false)
                    return null;
                FileStream fs = File.OpenRead(imagePath); //OpenRead
                int filelength = 0;
                filelength = (int)fs.Length; //获得文件长度
                Byte[] image = new Byte[filelength]; //建立一个字节数组
                fs.Read(image, 0, filelength); //按字节流读取
                System.Drawing.Image result = System.Drawing.Image.FromStream(fs);
                fs.Close();
                return new Bitmap(result);
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("读取图片异常!", "图片路径：" + imagePath + Environment.NewLine + ex);
                return null;
            }
        }

        [Description("根据图片名称或路径获取系统中图片")]
        public static BitmapImage GetBitmapImage(string imageName, System.Drawing.Size? size = null)
        {
            var image = GetImage(imageName);
            if (image == null)
                return null;
            if (size != null)
                image = image.ToImage(size.Value);
            var stream = new MemoryStream();
            image.Save(stream, ImageFormat.Png);
            BitmapImage bmp = new BitmapImage();
            bmp.BeginInit();
            bmp.StreamSource = stream;
            bmp.EndInit();
            return bmp;
        }

        public static BitmapImage GetBitmapImage(string imageName)
        {
            System.Drawing.Size? size = null;
            return GetBitmapImage(imageName, size);
        }

        [Description("获取全屏截图")]
        public static Image GetScreenImage()
        {
            try
            {
                //截取屏幕内容
                var width = Screen.AllScreens.Sum(s => s.Bounds.Width);
                var height = Screen.AllScreens.Max(s => s.Bounds.Height);
                System.Drawing.Size screen = new System.Drawing.Size(width, height);
                Bitmap memoryImage = new Bitmap(screen.Width, screen.Height);
                Graphics memoryGraphics = Graphics.FromImage(memoryImage);
                memoryGraphics.CopyFromScreen(0, 0, 0, 0, screen, CopyPixelOperation.SourceCopy);
                MemoryStream data = new MemoryStream();
                memoryImage.Save(data, ImageFormat.Png);
                return memoryImage;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        [Description("获取一个文件全路径（包括文件名和扩展名,返回null表示取消了选择，否则为全路径）" +
            "(fileName:文件名或文件路径，为null时将弹出对话框选择;" +
            "extension:文件的扩展名;createDirectory:指示目录不存在时是否创建目录)")]
        public static string GetFilePath(string fileName, string extension, bool createDirectory)
        {
            fileName = string.IsNullOrEmpty(fileName) ? null : fileName;
            var directory = Path.GetDirectoryName(fileName) ?? "";
            var name = Path.GetFileNameWithoutExtension(fileName) ?? "";
            extension = ((string.IsNullOrEmpty(extension) ? Path.GetExtension(fileName) : extension) ?? "").Trim('.');
            name = (string.IsNullOrEmpty(name) ? ("文件[" + DateTime.Now.ToString("yyyyMMddHHmmss") + "]") : name) + "." + extension;
            if (string.IsNullOrEmpty(directory))
            {
                var dig = new SaveFileDialog();
                dig.FileName = name;
                dig.Filter = "文件|*." + extension;
                dig.CheckFileExists = false;
                if (dig.ShowDialog() != DialogResult.OK)
                    return null;
                directory = Path.GetDirectoryName(dig.FileName);
                name = Path.GetFileName(dig.FileName);
            }
            if (createDirectory && Directory.Exists(directory) == false)
                Directory.CreateDirectory(directory);
            return Path.Combine(directory, name);
        }

        [Description("获取一个文件路径")]
        public static string GetFilePath(string fileName)
        {
            return GetFilePath(fileName, null, true);
        }

        [Description("获取一个文件夹目录")]
        public static string GetDirectory(string defalutDirectory = null)
        {
            var dig = new FolderBrowserDialog();
            dig.SelectedPath = defalutDirectory;
            if (dig.ShowDialog() == DialogResult.OK)
                return dig.SelectedPath;
            return null;
        }

        [Description("获取当前文件的下一个同名不同号的文件路径（如：新建文件夹1，文件1...等）")]
        public static string GetNextPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return "";
            var extent = Path.GetExtension(path);//后缀
            bool isdir = string.IsNullOrEmpty(extent);
            if ((isdir && !Directory.Exists(path)) || (!isdir && !File.Exists(path)))
                return path;
            var dir = Path.GetDirectoryName(path);//目录
            var filename = Path.GetFileNameWithoutExtension(path);//文件名称（不包括后缀）
            int index = 1;
            var newpath = path;
            while (true)
            {
                newpath = Path.Combine(dir, filename + "(" + index + ")" + extent);
                if ((isdir && Directory.Exists(newpath)) || (!isdir && File.Exists(newpath)))
                    index++;
                else
                    break;
            }
            return newpath;
        }

        [Description("获取指定文件夹下的文件(path:文件夹路径;filters:过滤条件;reverse:指示是否递归获取子集文件)")]
        public static List<string> GetFiles(string path, List<string> filters, bool reverse = false)
        {
            var result = new List<string>();
            if (Directory.Exists(path) == false)
                return result;
            filters = filters ?? new List<string>();
            if (filters.Count == 0)
                filters.Add("*");
            foreach (var filter in filters)
                result.AddRange(Directory.GetFiles(path, filter, reverse ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly));
            return result.Distinct().ToList();
        }

        #endregion 获取系统图标和路径

        #region 注册表操作

        [Description("读取注册表信息")]
        public static string ReadCurrentUserRegistryValue(string path, string keyName)
        {
            try
            {
                path = path.TrimStart('\\').Replace("HKEY_CURRENT_USER", "").TrimStart('\\');
                RegistryKey reg = Registry.CurrentUser.OpenSubKey(path);
                var value = reg?.GetValue(keyName)?.ToString() ?? "";
                if (reg != null)
                    reg.Close();
                return value;
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("读取注册表异常", "注册表路径：[" + path + "]中[" + keyName + "]" + Environment.NewLine + ex);
                return "";
            }
        }

        [Description("写注册表信息")]
        public static bool WriteCurrentUserRegistryValue(string path, string keyName, string value)
        {
            try
            {
                RegistryKey reg = Registry.CurrentUser.OpenSubKey(path, true);
                if (reg == null)
                    reg = Registry.CurrentUser.CreateSubKey(path);
                reg.OpenSubKey(keyName);
                reg.SetValue(keyName, value);
                reg.Close();
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("写入注册表异常", "注册表路径：" + path + "]中[" + keyName + "]的值【" + value + "】" + Environment.NewLine + ex);
                return false;
            }
        }

        [Description("判断注册表项是否存在")]
        public static bool IsCurrentUserRegistryExist(string path)
        {
            try
            {
                var reg = Registry.CurrentUser.OpenSubKey(path, false);
                if (reg == null)
                    return false;
                reg.Close();
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("注册表读取异常", "注册表路径：" + path + Environment.NewLine + ex);
                return false;
            }
        }

        [Description("判断注册表项是否存在")]
        public static bool IsCurrentUserRegistryValueExist(string path, string name)
        {
            try
            {
                var reg = Registry.CurrentUser.OpenSubKey(path, false);
                if (reg == null)
                    return false;
                var n = reg.GetValueNames().FirstOrDefault(s => s == name);
                reg.Close();
                return n != null;
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("注册表读取异常", "注册表路径：" + path + Environment.NewLine + ex);
                return false;
            }
        }

        [Description("删除注册表数据")]
        public static bool DeleteCurrentUserRegistry(string path)
        {
            try
            {
                Registry.CurrentUser.DeleteSubKey(path);
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("删除注册表异常", "注册表路径：" + path + Environment.NewLine + ex);
                return false;
            }
        }

        [Description("初始化WW插件")]
        public static void InitWWSoft()
        {
            try
            {
                var path = @"Software\WW\{F07847F9-0CED-4e2f-8771-9BEA3D1D10CC}";
                var value = ReadCurrentUserRegistryValue(path, "wmky");
                if (string.IsNullOrEmpty(path))
                    return;
                WriteCurrentUserRegistryValue(path, "wmky", "");
            }
            catch { }
        }

        #endregion 注册表操作

        [Description("获取本机的IP地址")]
        public static List<string> GetPCIP()
        {
            var entry = Dns.GetHostEntry(Dns.GetHostName());
            var result = entry?.AddressList?.Where(s => s.ToString().Split('.').Length == 4).Select(s => s.ToString()).ToList();
            return result ?? new List<string>();
        }

        [Description("获取本机的全名")]
        public static string GetPCName()
        {
            return Dns.GetHostEntry("localhost")?.HostName;
        }

        #region 文本文件读写

        [Description("读取文档内容")]
        public static string ReadFileContent(string path, bool decrypt = true)
        {
            try
            {
                if (File.Exists(path) == false)
                    return null;
                var content = File.ReadAllText(path, Encoding.UTF8);
                if (decrypt)
                    content = content.Decrypt();
                return content;
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("读取本地文件异常！", "读取本地文件异常:" + ex);
                return null;
            }
        }

        [Description("读取文件内容并转换为对象（decrypt:指示是否需要对文件内容解密）")]
        public static T ReadFileContent<T>(string path, bool decrypt = true, int maxDepth = 0)
        {
            var content = ReadFileContent(path, decrypt);
            if (content == null)
                return default(T);
            return content.ToObject<T>(maxDepth: maxDepth).Value;
        }

        [Description("将内容写入到指定文件（content：内容(可以是文本，也可以是对象)；encrypt：指示是否需要加密）")]
        public static bool WriteFileContent(string path, object content, bool encrypt = true)
        {
            try
            {
                var str = content is string ? content.ToString() : content?.ToJson();
                if (encrypt)
                    str = str?.Encrypt();
                var dir = Path.GetDirectoryName(path);
                if (Directory.Exists(dir) == false)
                    Directory.CreateDirectory(dir);
                File.WriteAllText(path, str, Encoding.UTF8);
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("写入本地文本异常", ex);
                return false;
            }
        }

        #endregion 文本文件读写

        #region 反射dll

        [Description("反射某名称的dll（在当前集成的所有搜索目录下查找）(assemblyName:dll名称（不是全路径）)")]
        public static Result<Assembly> CreateAssemblyByName(string assemblyName)
        {
            try
            {
                var path = Path.Combine(DataLibrary.WorkDirectory, assemblyName);
                if (File.Exists(path) == false)
                    return new Result<Assembly>(false, null, "程序集不存在！");
                var dll = Assembly.LoadFrom(path);
                return new Result<Assembly>(true, dll);
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("反射生成程序集异常", ex);
                return new Result<Assembly>(false, null, ex.Message);
            }
        }

        [Description("反射某名称的dll(assemblyPath:dll全路径)")]
        public static Result<Assembly> CreateAssemblyByPath(string assemblyPath)
        {
            try
            {
                var dll = Assembly.LoadFrom(assemblyPath);
                return new Result<Assembly>(true, dll);
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("反射生成程序集异常", ex);
                return new Result<Assembly>(false, null, ex.Message);
            }
        }

        [Description("反射某名称dll中某名称的类对象(className:类名称（注意：类名称为当前类的命名空间名加类名）)")]
        public static Result<T> CreateClassObject<T>(Assembly assembly, string className) where T : class
        {
            try
            {
                var instance = assembly.CreateInstance(className) as T;
                if (instance == null)
                    return new Result<T>(false, null, "不能创建名称为【" + className + "】的类对象！");
                return new Result<T>(true, instance);
            }
            catch (Exception ex)
            {
                LogHelper.WriteExceptionLog("创建名称为【" + className + "】的类对象异常", ex);
                return new Result<T>(false, null, ex.Message);
            }
        }

        [Description("反射某名称dll中某名称的类对象(assembly:dll名称（不是全路径）;className:类名称（注意：类名称为当前类的命名空间名加类名）)")]
        public static Result<T> CreateClassObject<T>(string assemblyName, string className) where T : class
        {
            var dll = CreateAssemblyByName(assemblyName);
            if (dll.IsSuccess == false)
                return new Result<T>(false, null, dll.Message);
            return CreateClassObject<T>(dll.Value, className);
        }

        #endregion 反射dll
    }
}