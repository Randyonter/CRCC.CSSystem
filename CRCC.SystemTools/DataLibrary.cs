/****************************************************************************************************************************************************************
 *
 * 版权所有: 2023  All Rights Reserved.
 * 文 件 名:  DataLibrary
 * 描    述:  系统参数库
 * 创 建 者：  Ltf
 * 创建时间：  2023/3/29 15:43:58
 * 历史更新记录:
 * =============================================================================================
*修改标记
*修改时间：2023/3/29 15:43:58
*修改人： Ltf
*版本号： V1.0.0.0
*描述：
 * =============================================================================================

****************************************************************************************************************************************************************/

using System.ComponentModel;
using System.IO;
using System.Reflection;

namespace CRCC.SystemTools
{
    public static class DataLibrary
    {
        #region 目录

        private static string workDirectory;

        [Description("获取程序目录")]
        public static string WorkDirectory
        {
            get
            {
                if (string.IsNullOrEmpty(workDirectory))
                    workDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase.Substring(8));
                return workDirectory;
            }
        }

        [Description("获取图标目录全路径")]
        public static string ImagesDirectory
        { get { string path = Path.Combine(WorkDirectory, "Images"); if (Directory.Exists(path) == false) Directory.CreateDirectory(path); return path; } }

        [Description("获取日志目录全路径")]
        public static string LogDirectory
        { get { string path = Path.Combine(WorkDirectory, "Log"); if (Directory.Exists(path) == false) Directory.CreateDirectory(path); return path; } }

        [Description("获取配置目录全路径")]
        public static string ConfigDirectory
        { get { string path = Path.Combine(WorkDirectory, "Config"); if (Directory.Exists(path) == false) Directory.CreateDirectory(path); return path; } }

        [Description("获取模板目录全路径")]
        public static string TemplateDirectory
        { get { string path = Path.Combine(WorkDirectory, "Template"); if (Directory.Exists(path) == false) Directory.CreateDirectory(path); return path; } }

        [Description("获取临时文件目录全路径")]
        public static string TempDirectory
        { get { string path = Path.Combine(WorkDirectory, "Temp"); if (Directory.Exists(path) == false) Directory.CreateDirectory(path); return path; } }

        [Description("获取帮助文件目录全路径")]
        public static string HelpDirectory
        { get { string path = Path.Combine(WorkDirectory, "Help"); if (Directory.Exists(path) == false) Directory.CreateDirectory(path); return path; } }

        [Description("获取更新文件目录全路径")]
        public static string UpdateDirectory
        { get { string path = Path.Combine(WorkDirectory, "Update"); if (Directory.Exists(path) == false) Directory.CreateDirectory(path); return path; } }

        [Description("获取本地卷目录全路径")]
        public static string LocalVolumnDirectory
        { get { string path = Path.Combine(WorkDirectory, "Innervolumn"); if (Directory.Exists(path) == false) Directory.CreateDirectory(path); return path; } }

        [Description("获取缓存目录全路径")]
        public static string CacheDirectory
        { get { string path = Path.Combine(WorkDirectory, "Cache"); if (Directory.Exists(path) == false) Directory.CreateDirectory(path); return path; } }

        #endregion 目录

        public static readonly double DefaultMinPercent = 0.99;
    }
}