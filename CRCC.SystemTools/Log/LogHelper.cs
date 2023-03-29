/****************************************************************************************************************************************************************
 *
 * 版权所有: 2023  All Rights Reserved.
 * 文 件 名:  LogHelper
 * 描    述:
 * 创 建 者：  Ltf
 * 创建时间：  2023/3/29 15:40:10
 * 历史更新记录:
 * =============================================================================================
*修改标记
*修改时间：2023/3/29 15:40:10
*修改人： Ltf
*版本号： V1.0.0.0
*描述：
 * =============================================================================================

****************************************************************************************************************************************************************/

using System;
using System.ComponentModel;
using System.IO;
using System.Text;

namespace CRCC.SystemTools.Log
{
    public class LogHelper
    {
        [Description("输出日志")]
        public static void WriteLog(string title, LogType logType, object content)
        {
            if (content == null)
                return;
            try
            {
                var path = Path.Combine(DataLibrary.LogDirectory, "[" + DateTime.Now.ToString("yyyy-MM-dd") + "]" + logType + ".log");
                var log = "************************【" + DateTime.Now.ToString() + "】*****************************";
                log += Environment.NewLine + content + Environment.NewLine;
                File.AppendAllText(path, log, Encoding.UTF8);
            }
            catch { }
        }

        [Description("输出异常")]
        public static void WriteExceptionLog(string title, object content)
        {
            WriteLog(title, LogType.Exception, content);
        }

        [Description("输出调试")]
        public static void WriteDebugLog(string title, object content)
        {
            WriteLog(title, LogType.Debug, content);
        }
    }

    [Description("日志类型")]
    public enum LogType
    {
        [Description("异常日志")]
        Exception = 0,

        [Description("调试日志")]
        Debug = 1,

        [Description("操作日志")]
        Operate = 2,
    }
}