/****************************************************************************************************************************************************************
 *
 * 版权所有: 2023  All Rights Reserved.
 * 文 件 名:  ExpandHelper
 * 描    述:
 * 创 建 者：  Ltf
 * 创建时间：  2023/3/29 15:48:02
 * 历史更新记录:
 * =============================================================================================
*修改标记
*修改时间：2023/3/29 15:48:02
*修改人： Ltf
*版本号： V1.0.0.0
*描述：
 * =============================================================================================

****************************************************************************************************************************************************************/

using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace CRCC.SystemTools
{
    public static class ExpandHelper
    {
        [Description("将对象序列化为字符串")]
        public static string ToJson(this object obj)
        {
            var st = new JsonSerializerSettings
            {
                NullValueHandling = NullValueHandling.Ignore,
                MissingMemberHandling = MissingMemberHandling.Ignore
            };
            return JsonConvert.SerializeObject(obj, st);
        }

        [Description("强制转换数据对象类型，转换失败返回所转类型的默认值")]
        public static Result<T> ToValue<T>(this object value, T defaultValue = default(T))
        {
            if (value == null)
                return new Result<T>(false, defaultValue, "数据为空！");
            var type = typeof(T);
            if (type.IsEnum)
                return new Result<T>(true, value.ToString().ToEnum<T>(defaultValue), null);
            if (value.GetType().Equals(type))
                return new Result<T>(true, (T)value, null);
            if (type == typeof(object))
                return new Result<T>(true, (T)value);
            bool result = false;
            object resultValue = defaultValue;
            if (type == typeof(string))
            {
                result = true;
                resultValue = value.ToString();
            }
            else if (type == typeof(int))
            {
                int temp;
                result = int.TryParse(value.ToString(), out temp);
                resultValue = temp;
            }
            else if (type == typeof(double))
            {
                double temp;
                result = double.TryParse(value.ToString(), out temp);
                resultValue = temp;
            }
            else if (type == typeof(bool))
            {
                bool temp;
                result = bool.TryParse(value.ToString(), out temp);
                resultValue = temp;
            }
            else if (type == typeof(float))
            {
                float temp;
                result = float.TryParse(value.ToString(), out temp);
                resultValue = temp;
            }
            else if (type == typeof(decimal))
            {
                decimal temp;
                result = decimal.TryParse(value.ToString(), out temp);
                resultValue = temp;
            }
            else if (type == typeof(DateTime))
            {
                DateTime temp = DateTime.Now;
                result = DateTime.TryParse(value.ToString(), out temp);
                resultValue = temp;
            }
            else if (type == typeof(Color))
            {
                int temp;
                result = int.TryParse(value.ToString(), out temp);
                resultValue = Color.FromArgb(temp);
            }
            else
            {
                try
                {
                    resultValue = value.ToString().ToObject<T>().Value;
                    result = true;
                }
                catch
                {
                    result = false;
                }
            }
            return new Result<T>(result, result ? (T)resultValue : defaultValue, null);
        }

        [Description("强制转换某类型的数据值")]
        public static object ToValue(this object value, Type type)
        {
            if (value == null)
                return null;
            if (type.IsEnum)
            {
                if (string.IsNullOrEmpty(value?.ToString()))
                    return Enum.Parse(type, Enum.GetNames(type).FirstOrDefault());
                return Enum.Parse(type, value?.ToString());
            }
            if (value.GetType() == type)
                return value;
            else if (type == typeof(string))
                return value.ToValue<string>().Value;
            else if (type == typeof(int))
                return value.ToValue<int>().Value;
            else if (type == typeof(double))
                return value.ToValue<double>().Value;
            else if (type == typeof(bool))
                return value.ToValue<bool>().Value;
            else if (type == typeof(float))
                return value.ToValue<float>().Value;
            else if (type == typeof(decimal))
                return value.ToValue<decimal>().Value;
            else if (type == typeof(DateTime))
                return value.ToValue<DateTime>().Value;
            return value;
        }

        [Description("获取枚举值的属性描述")]
        public static string GetDescription(this Enum obj)
        {
            try
            {
                if (obj == null)
                    return string.Empty;
                Type _enumType = obj.GetType();
                var fi = _enumType.GetField(System.Enum.GetName(_enumType, obj));
                var dna = (DescriptionAttribute)Attribute.GetCustomAttribute(fi, typeof(DescriptionAttribute));
                if (dna != null && !string.IsNullOrEmpty(dna.Description))
                    return dna.Description;
                return obj.ToString();
            }
            catch { return obj.ToString(); }
        }

        [Description("将图片转换为指定大小的图片")]
        public static Image ToImage(this Image srcImage, Size newsize)
        {
            if (srcImage == null || newsize == null)
                return srcImage;
            return new Bitmap(srcImage, newsize);
        }

        [Description("获取当前类型的所有基类名称")]
        public static List<string> GetBaseTypeNames(this Type type)
        {
            var types = new List<string>();
            if (type == null)
                return types;
            var basetype = type;
            while (basetype != null)
            {
                types.Add(basetype.Name);
                basetype = basetype.BaseType;
            }
            return types;
        }

        [Description("判断当前类型是否继承自指定的类型")]
        public static bool InheritOf(this Type type, Type baseType, bool includeSelf = true)
        {
            if (type == null || baseType == null)
                return false;
            if (baseType.IsInterface)
            {
                var interfacevalue = type.GetInterface(baseType.Name);
                return interfacevalue != null;
            }
            Type basetype = includeSelf ? type : type?.BaseType;
            while (basetype != null)
            {
                if (basetype == baseType)
                    return true;
                basetype = basetype.BaseType;
            }
            return false;
        }

        #region 对字符串（string）的扩展

        [Description("将字符串转换为整数")]
        public static int ToInt(this string str)
        {
            int result = 0;
            if (int.TryParse(str, out result) == false)
                result = 0;
            return result;
        }

        [Description("将字符串转换为指定位数的小数(decimals:保留的小数位数)")]
        public static float ToFloat(this string str, int decimals = 2)
        {
            float result = 0;
            if (float.TryParse(str, out result) == false)
                result = 0;
            return (float)Math.Round(result, decimals);
        }

        [Description("将字符串转换为指定位数的小数(decimals:保留的小数位数)")]
        public static double ToDouble(this string str, int decimals = 2)
        {
            double result = 0;
            if (double.TryParse(str, out result) == false)
                result = 0;
            return Math.Round(result, decimals);
        }

        [Description("将字符串转换为指定位数的小数(decimals:保留的小数位数)")]
        public static decimal ToDecimal(this string str, int decimals = 2)
        {
            decimal result = 0;
            if (decimal.TryParse(str, out result) == false)
                result = 0;
            return Math.Round(result, decimals);
        }

        [Description("将字符串转换为Bool值")]
        public static bool ToBool(this string str)
        {
            return str?.ToLower() == "true" || str == "是";
        }

        [Description("将字符串转换为时间")]
        public static DateTime ToTime(this string str)
        {
            var result = DateTime.Now;
            if (DateTime.TryParse(str, out result) == false)
                result = DateTime.Now;
            return result;
        }

        [Description("将字符串转换为日期")]
        public static DateTime ToDate(this string str)
        {
            DateTime result = DateTime.Now;
            if (DateTime.TryParse(str, out result) == false)
                result = DateTime.Now;
            return result.Date;
        }

        [Description("将字符串转换为枚举值")]
        public static T ToEnum<T>(this string str, T defaultvalue = default(T))
        {
            try
            {
                return (T)Enum.Parse(typeof(T), str);
            }
            catch
            {
                return defaultvalue;
            }
        }

        [Description("根据描述获取对应的枚举项")]
        public static T ToEnumByDescription<T>(this string str, T defaultvalue = default(T))
        {
            return CommonHelper.GetEnumByDescription(str, defaultvalue);
        }

        [Description("通过反序列化将字符串转换为对象")]
        public static Result<T> ToObject<T>(this string str, T defaultValue = default(T), int maxDepth = 0)
        {
            try
            {
                var set = new JsonSerializerSettings();
                set.NullValueHandling = NullValueHandling.Ignore;
                set.ObjectCreationHandling = ObjectCreationHandling.Auto;
                if (maxDepth > 0)
                    set.MaxDepth = maxDepth;
                return new Result<T>(true, JsonConvert.DeserializeObject<T>(str, set));
            }
            catch (Exception ex) { return new Result<T>(false, defaultValue, ex.Message); }
        }

        [Description("获取一段字符的尺寸（以像素为单位）")]
        public static Size GetSize(this string text, Font font)
        {
            font = font ?? new Font("宋体", 9);
            if (string.IsNullOrEmpty(text))
                return new Size() { Width = 0, Height = 0 };
            return TextRenderer.MeasureText(text, font);
        }

        private static byte[] makeMD5(byte[] original)
        {
            MD5CryptoServiceProvider hashmd5 = new MD5CryptoServiceProvider();
            byte[] keyhash = hashmd5.ComputeHash(original);
            hashmd5 = null;
            return keyhash;
        }

        [Description("字符串加密")]
        public static string Encrypt(this string str)
        {
            try
            {
                byte[] buff = Encoding.Default.GetBytes(str);
                byte[] kb = Encoding.Default.GetBytes("zjc");
                TripleDESCryptoServiceProvider des = new TripleDESCryptoServiceProvider();
                des.Key = makeMD5(kb);
                des.Mode = CipherMode.ECB;
                var byes = des.CreateEncryptor().TransformFinalBlock(buff, 0, buff.Length);
                return Convert.ToBase64String(byes);
            }
            catch
            {
                return "";
            }
        }

        [Description("字符串解密")]
        public static string Decrypt(this string str)
        {
            try
            {
                if (string.IsNullOrEmpty(str))
                    return "";
                byte[] buff = Convert.FromBase64String(str);
                byte[] kb = System.Text.Encoding.Default.GetBytes("zjc");
                TripleDESCryptoServiceProvider des = new TripleDESCryptoServiceProvider();
                des.Key = makeMD5(kb);
                des.Mode = CipherMode.ECB;
                var byes = des.CreateDecryptor().TransformFinalBlock(buff, 0, buff.Length);
                return Encoding.Default.GetString(byes);
            }
            catch
            {
                return "";
            }
        }

        [Description("判断一个值是否符合正则表达式")]
        public static bool IsMatch(this string str, string regStr)
        {
            return new Regex(regStr).IsMatch(str);
        }

        [Description("判断是否包含数字")]
        public static bool ContainNumber(this string str)
        {
            var v = str ?? "";
            return Regex.IsMatch(v, @"\d+");
        }

        [Description("判断是否为纯数字")]
        public static bool IsNumber(this string str)
        {
            double d;
            return double.TryParse(str, out d);
        }

        [Description("判断是否为纯字母(不区分大小写)")]
        public static bool IsAZChar(this string str)
        {
            var regstr = @"^[A-Za-z]+$";
            return IsMatch(str, regstr);
        }

        [Description("判断当前字符串是否为时间格式")]
        public static bool IsTime(this string str)
        {
            DateTime time;
            return DateTime.TryParse(str, out time);
        }

        [Description("判断当前字符串是否为布尔值")]
        public static bool IsBool(this string str)
        {
            bool result;
            return bool.TryParse(str, out result);
        }

        [Description("判断当前字符串是否为四则运算操作符")]
        public static bool IsOperate(this string str)
        {
            return str == "+" || str == "-" || str == "*" || str == "/" || str == "%";
        }

        [Description("提取文本里面的数字（包括正数，负数，小数）")]
        public static List<double> GetNumber(this string str)
        {
            if (string.IsNullOrEmpty(str))
                return new List<double>();
            var par = @"[+-]?\d+[\.]?\d*";
            return Regex.Matches(str, par).Cast<Match>().Select(s => double.Parse(s.Value)).ToList();
        }

        [Description("获取两个字符串的相似度(0-1)(srcStr:源字符串;desStr:目标字符串;sameKey:相似内容的权重1-10)")]
        public static double GetSimilarity(this string srcStr, string desStr, int sameKey = 5)
        {
            if (string.IsNullOrEmpty(srcStr) || string.IsNullOrEmpty(desStr))
                return 0;
            var sm = srcStr.Intersect(desStr).ToList();//两字符串字符的交集
            var str1sm = sm.ToDictionary(s => s, s => srcStr.Count(m => m == s));//第一个字符串中交集部分的数量
            var str2sm = sm.ToDictionary(s => s, s => desStr.Count(m => m == s));//第二个字符串中交集部分的数量
            var smcount = sm.Count == 0 ? 0 : sm.Sum(s => str1sm[s] > str2sm[s] ? str2sm[s] : str1sm[s]);//交集字符的数量
            return CommonHelper.GetSimilarityByCount(srcStr.Length, desStr.Length, smcount, sameKey);
        }

        [Description("根据指定字符，分割字符串")]
        public static List<string> Split(this string src, string sp)
        {
            src = src ?? "";
            if (string.IsNullOrEmpty(sp))
                return new List<string> { src };
            return src.Split(new string[] { sp }, StringSplitOptions.RemoveEmptyEntries).ToList();
        }

        [Description("获取被匹配的值")]
        public static List<string> GetMatchValue(this string str, string regStr)
        {
            return new Regex(regStr).Matches(str).Cast<Match>().Select(s => s.Value).ToList();
        }

        [Description("替换指定字符里面的指定字符")]
        public static string Replace(this string srcStr, List<string> replaceStr, string desStr)
        {
            if (string.IsNullOrEmpty(srcStr) || replaceStr == null || replaceStr.Count == 0)
                return srcStr;
            foreach (var sp in replaceStr)
                srcStr = srcStr.Replace(sp, desStr);
            return srcStr;
        }

        #endregion 对字符串（string）的扩展

        #region 对数字的扩展

        [Description("转换为整数（四舍五入）")]
        public static int ToInt(this float value)
        {
            return (int)Math.Round(value, 0);
        }

        [Description("转换为整数（四舍五入）")]
        public static int ToInt(this double value)
        {
            return (int)Math.Round(value, 0);
        }

        [Description("转换为整数（四舍五入）")]
        public static int ToInt(this decimal value)
        {
            return (int)Math.Round(value, 0);
        }

        [Description("转换为整数")]
        public static int ToInt(this byte[] value)
        {
            //低位在前，高位在后
            const int offset = 0;
            return (int)((value[offset] & 0xFF) |
                              ((value[offset + 1] << 8) & 0xFF00) |
                              ((value[offset + 2] << 16) & 0xFF0000) |
                              ((value[offset + 3] << 24) & 0xFF000000));
        }

        [Description("转换为指定位数的浮点数")]
        public static float ToFloat(this float value, int decimals = 2)
        {
            return (float)Math.Round(value, decimals);
        }

        [Description("转换为指定位数的浮点数")]
        public static float ToFloat(this double value, int decimals = 2)
        {
            return (float)Math.Round(value, decimals);
        }

        [Description("转换为指定位数的浮点数")]
        public static float ToFloat(this decimal value, int decimals = 2)
        {
            return (float)Math.Round(value, decimals);
        }

        [Description("转换为对应的字节")]
        public static byte[] ToBytes(this int value)
        {
            var byteSrc = new byte[4];
            byteSrc[3] = (byte)((value & 0xFF000000) >> 24);
            byteSrc[2] = (byte)((value & 0x00FF0000) >> 16);
            byteSrc[1] = (byte)((value & 0x0000FF00) >> 8);
            byteSrc[0] = (byte)((value & 0x000000FF));
            return byteSrc;
        }

        [Description("转换为指定位数的十进制数")]
        public static decimal ToDecimal(this float value, int decimals = 2)
        {
            return (decimal)Math.Round(value, decimals);
        }

        [Description("转换为指定位数的十进制数")]
        public static decimal ToDecimal(this double value, int decimals = 2)
        {
            return (decimal)Math.Round(value, decimals);
        }

        [Description("将弧度转换为角度")]
        public static double ToAngle(this double radian)
        {
            return (180 * radian) / Math.PI;
        }

        [Description("将角度转换为弧度")]
        public static double ToRadian(this double angle)
        {
            return angle * Math.PI / 180;
        }

        [Description("获取当前角度关于X轴的镜像角度")]
        public static double GetXMirrorAngle(this double radian)
        {
            var pi2 = 2 * Math.PI;
            return pi2 - radian % pi2;
        }

        [Description("获取当前角度关于Y轴的镜像角度")]
        public static double GetYMirrorAngle(this double radian)
        {
            var pi2 = 2 * Math.PI;
            return (pi2 + (Math.PI - radian % pi2)) % pi2;
        }

        [Description("判断一个弧度是否为直角（如：0°，90°，180°，270°，360°）")]
        public static bool IsRightAngle(this double radian, double minPercent = -1)
        {
            radian = radian % (2 * Math.PI);
            return radian.EqualsWith(0, minPercent) || radian.EqualsWith(Math.PI / 2, minPercent) || radian.EqualsWith(Math.PI, minPercent)
                || radian.EqualsWith(3 * Math.PI / 2, minPercent) || radian.EqualsWith(2 * Math.PI, minPercent);
        }

        [Description("将弧度在一定误差范围内格式化为整角（0，90，270）的弧度制")]
        public static double FormatRadian(this double radian, double rang = 0.05)
        {
            var p = Math.PI;
            rang = rang < 0 ? 0 : rang > p ? p : rang;
            var angle = radian % (2 * Math.PI);
            var isfs = angle < 0;
            angle = isfs ? -angle : angle;
            var result = Math.Abs(angle - 0) <= rang || Math.Abs(angle - 2 * p) <= rang ? 0 : Math.Abs(angle - p / 2) < rang ? p / 2 : Math.Abs(angle - p) < rang ? p : Math.Abs(angle - p * 3 / 2) < rang ? p * 3 / 2 : angle;
            return isfs ? -result : result;
        }

        [Description("将字节长度转换为文件大小（B,KB,MB,G）")]
        public static string ToFileSize(this long length)
        {
            var kbSize = length / 1024;
            if (kbSize < 1)
                string.Format("{0}B", length.ToString("N2"));
            var mbSize = kbSize / 1024;
            if (mbSize < 1)
                return string.Format("{0}KB", kbSize.ToString("N2"));
            var gbSize = mbSize / 1024;
            if (gbSize < 1)
                return string.Format("{0}MB", mbSize.ToString("N2"));
            return string.Format("{0}GB", gbSize.ToString("N2"));
        }

        [Description("获取两个数的相似性(0-1)，如果参考点在两数之间，则相似度为0")]
        public static double GetSimilarity(this double srcValue, double desValue, double referValue = 0)
        {
            if (Math.Abs(srcValue - desValue) < 0.001)
                return 1;
            if ((referValue >= srcValue && referValue <= desValue) || (referValue <= srcValue && referValue >= desValue))
                return 0;
            var d1 = Math.Abs(srcValue - referValue);
            var d2 = Math.Abs(desValue - referValue);
            var min = Math.Min(d1, d2); var max = Math.Max(d1, d2);
            var result = min / max;
            return result < 0.000001 ? 0 : result;
        }

        [Description("比较两个值在指定相似度内的大小（返回0表示相同，返回1表示当前值比第二个值大，返回-1表示当前值比第二个值小）")]
        public static int CompareWith(this double srcValue, double desValue, double minPercent = -1, double referValue = 0)
        {
            minPercent = minPercent <= 0 || minPercent > 1 ? DataLibrary.DefaultMinPercent : minPercent;
            var result = srcValue.GetSimilarity(desValue, referValue);
            if (result > 0)
                return result >= minPercent ? 0 : srcValue > desValue ? 1 : -1;
            var min = Math.Min(srcValue, desValue); var max = Math.Max(srcValue, desValue);
            var error = max < 1 ? 0.0001 : max < 10 ? 0.001 : max < 100 ? 0.01 : max < 1000 ? 0.1 : max < 10000 ? 10 : 50;
            return max - min <= error ? 0 : srcValue > desValue ? 1 : -1;
        }

        [Description("判断两个值在指定相似度内是否相等")]
        public static bool EqualsWith(this double srcValue, double desValue, double minPercent = -1, double referValue = 0)
        {
            return CompareWith(srcValue, desValue, minPercent, referValue) == 0;
        }

        #endregion 对数字的扩展

        #region 对DataTable的扩展

        [Description("将DataTable中数据转换为对象")]
        public static List<T> ToObjects<T>(this DataTable table)
        {
            return table?.ToJson()?.ToObject<List<T>>().Value;
        }

        [Description("将一行数据转换为对象")]
        public static T ToObject<T>(this DataRow row)
        {
            if (row == null)
                return default(T);
            return row.ToJson().ToObject<T>().Value;
        }

        [Description("将一行数据转换为列表")]
        public static Dictionary<string, object> ToDictionary(this DataRow row)
        {
            if (row == null)
                return null;
            return row.Table.Columns.Cast<DataColumn>().ToDictionary(s => s.ColumnName, s => row[s.ColumnName]);
        }

        [Description("将当前行DataTable转换为列表")]
        public static List<Dictionary<string, object>> ToList(this DataTable table)
        {
            if (table == null)
                return null;
            var columns = table.Columns.Cast<DataColumn>().Select(s => s.ColumnName).ToList();
            return table.Rows.Cast<DataRow>().Select(s => columns.ToDictionary(m => m, n => s[n])).ToList();
        }

        #endregion 对DataTable的扩展

        #region 对集合的扩展

        [Description("将多个列表中的值合并到一个新字集合中")]
        public static List<T> Merge<T>(this IEnumerable<IEnumerable<T>> lists, bool distinct = false)
        {
            if (lists == null)
                return null;
            var list = new List<T>();
            foreach (var l in lists)
            {
                if (l != null)
                    list.AddRange(l);
            }
            return distinct ? list.Distinct().ToList() : list;
        }

        [Description("将多个字典中的值合并到一个新字典中")]
        public static Dictionary<T1, T2> Merge<T1, T2>(this IEnumerable<Dictionary<T1, T2>> dics)
        {
            if (dics == null)
                return null;
            var dic = new Dictionary<T1, T2>();
            foreach (var list in dics)
            {
                if (list == null)
                    continue;
                foreach (var key in list)
                {
                    if (dic.ContainsKey(key.Key) == false)
                        dic.Add(key.Key, key.Value);
                    else
                        dic[key.Key] = key.Value;
                }
            }
            return dic;
        }

        [Description("将多个字典中的值合并到一个新字典中")]
        public static Dictionary<T1, List<T2>> Merges<T1, T2>(this IEnumerable<Dictionary<T1, T2>> dics)
        {
            if (dics == null)
                return null;
            var dic = new Dictionary<T1, List<T2>>();
            foreach (var list in dics)
            {
                if (list == null)
                    continue;
                foreach (var key in list)
                {
                    if (dic.ContainsKey(key.Key) == false)
                        dic.Add(key.Key, new List<T2>() { key.Value });
                    else
                        dic[key.Key].Add(key.Value);
                }
            }
            return dic;
        }

        [Description("将一系列值添加到当前字典中")]
        public static void AddRange<T1, T2>(this Dictionary<T1, T2> source, IEnumerable<KeyValuePair<T1, T2>> items)
        {
            if (source == null || items == null || items.Any() == false)
                return;
            foreach (var item in items)
            {
                if (source.ContainsKey(item.Key))
                    continue;
                source.Add(item.Key, item.Value);
            }
        }

        [Description("通过自定义逻辑判断指定集合中是否包含某对象")]
        public static bool Contains<T>(this IEnumerable<T> items, Func<T, bool> fun)
        {
            if (items == null || items.Any() == false || fun == null)
                return false;
            foreach (var item in items)
            {
                var result = fun.Invoke(item);
                if (result)
                    return true;
            }
            return false;
        }

        [Description("将一个列表集合转换为Item表")]
        public static DataTable ToItemTable(this Dictionary<string, string> values, bool addEmpty = false)
        {
            var tb = CommonHelper.CreateItemTable(addEmpty);
            foreach (var value in values)
            {
                if (addEmpty && string.IsNullOrEmpty(value.Key))
                    continue;
                var datarow = tb.NewRow();
                datarow["Key"] = value.Key;
                datarow["Value"] = value.Value;
                tb.Rows.Add(datarow);
            }
            return tb;
        }

        [Description("将一个列表集合转换为Item表")]
        public static DataTable ToItemTable(this List<string> values, bool addEmpty = false)
        {
            var tb = CommonHelper.CreateItemTable(addEmpty);
            foreach (var value in values)
            {
                var datarow = tb.NewRow();
                datarow["Key"] = datarow["Value"] = value;
                tb.Rows.Add(datarow);
            }
            return tb;
        }

        [Description("获取的字典中值")]
        public static T2 GetValue<T1, T2>(this Dictionary<T1, T2> dic, T1 key, T2 defaultValue = default(T2))
        {
            return dic.ContainsKey(key) ? dic[key] : defaultValue;
        }

        #endregion 对集合的扩展

        #region 对时间的扩展

        [Description("将时间转换为格式为【yyyy-MM-dd】的字符")]
        public static string ToDateString(this DateTime time)
        {
            return time.ToString("yyyy-MM-dd");
        }

        [Description("将时间转换为格式为【yyyy-MM-dd HH:mm:ss】的字符")]
        public static string ToTimeString(this DateTime time)
        {
            return time.ToString("yyyy-MM-dd HH:mm:ss");
        }

        [Description("转换为星期")]
        public static string ToWeek(this DayOfWeek week)
        {
            switch (week)
            {
                case DayOfWeek.Monday:
                    return "星期一";

                case DayOfWeek.Tuesday:
                    return "星期二";

                case DayOfWeek.Wednesday:
                    return "星期三";

                case DayOfWeek.Thursday:
                    return "星期四";

                case DayOfWeek.Friday:
                    return "星期五";

                case DayOfWeek.Saturday:
                    return "星期六";

                case DayOfWeek.Sunday:
                    return "星期日";

                default:
                    return "不知道";
            }
        }

        [Description("获取从指定时间开始，到下个星期几的时间")]
        public static DateTime GetNextWeekTime(this DateTime startTime, DayOfWeek nextWeek)
        {
            var curweek = startTime.DayOfWeek;
            var bdays = nextWeek == DayOfWeek.Sunday && curweek == DayOfWeek.Saturday ? 1 : nextWeek == curweek ? 7 : nextWeek > curweek ? nextWeek - curweek : 7 + nextWeek - curweek;
            return startTime.AddDays(bdays);
        }

        #endregion 对时间的扩展

        [Description("根据列号和数量拆分表(columnStep:每个拆分项保留的列数；startIndex:允许被拆分列的起始索引号；endIndex:允许被拆分列的终止索引号；)")]
        public static List<DataTable> SplitByColumn(this DataTable srcTable, int columnStep, int startIndex = 0, int endIndex = 0)
        {
            if (srcTable == null)
                return new List<DataTable>();
            if (columnStep <= 0)
                return new List<DataTable>() { srcTable };
            var cCount = srcTable.Columns.Count;
            startIndex = startIndex < 0 ? 0 : (startIndex >= cCount ? (cCount - 1) : startIndex);//所有被拆分列的起始索引号
            endIndex = endIndex <= startIndex || endIndex >= cCount ? (cCount - 1) : endIndex;//所有被拆分列的结束索引号
            var splitcount = endIndex - startIndex + 1;
            var tbcount = (splitcount / columnStep) + ((splitcount % columnStep) == 0 ? 0 : 1);//根据传入的每个表列数计算需要拆分成几个表
            var tbs = new List<DataTable>();
            for (int i = 0; i < tbcount; i++)//根据数量拆分原表
            {
                var newtb = srcTable.Copy();
                var startindex = startIndex + (i * columnStep);//当前表的起始索引号
                var endindex = startindex + columnStep - 1;//当前表的结束索引号
                for (int j = endIndex; j >= startIndex; j--)
                {
                    if (j >= startindex && j <= endindex)
                        continue;
                    newtb.Columns.RemoveAt(j);
                }
                tbs.Add(newtb);
            }
            return tbs;
        }

        [Description("根据行号和数量拆分表(rowStep:每个拆分项保留的行数；startIndex:允许被拆分行的起始索引号；endIndex:允许被拆分行的终止索引号；)")]
        public static List<DataTable> SplitByRow(this DataTable srcTable, int rowStep, int startIndex = 0, int endIndex = 0)
        {
            if (srcTable == null)
                return new List<DataTable>();
            if (rowStep <= 0)
                return new List<DataTable>() { srcTable };
            var rCount = srcTable.Rows.Count;
            startIndex = startIndex < 0 ? 0 : (startIndex >= rCount ? (rCount - 1) : startIndex);//所有被拆分行的起始索引号
            endIndex = endIndex <= startIndex || endIndex >= rCount ? (rCount - 1) : endIndex;//所有被拆分行的结束索引号
            var splitcount = endIndex - startIndex + 1;
            var tbcount = (splitcount / rowStep) + ((splitcount % rowStep) == 0 ? 0 : 1);//根据传入的每个表行数计算需要拆分成几个表
            var tbs = new List<DataTable>();
            for (int i = 0; i < tbcount; i++)//根据数量拆分原表
            {
                var newtb = srcTable.Copy();
                var startindex = startIndex + (i * rowStep);//当前表的起始索引号
                var endindex = startindex + rowStep - 1;//当前表的结束索引号
                for (int j = endIndex; j >= startIndex; j--)
                {
                    if (j >= startindex && j <= endindex)
                        continue;
                    newtb.Rows.RemoveAt(j);
                }
                tbs.Add(newtb);
            }
            return tbs;
        }
    }
}