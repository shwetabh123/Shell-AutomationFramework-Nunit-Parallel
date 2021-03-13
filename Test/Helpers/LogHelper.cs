using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace Test.Helpers
{
   public static class LogHelper
    {
        public static StreamWriter _streamw = null;
        public static string partiallogFilePath = "\\" + "Log-" + System.DateTime.Now.ToString("MM-dd-yyyy_HHmmss") + "." + "txt";
        static LogHelper()
        {
        }
        public static void WriteLog( String filepath, String logmessage)
        {
            StreamWriter log;
            FileStream fileStream = null;
            DirectoryInfo logDirInfo = null;
            FileInfo logFileInfo;
            string logFilePath = filepath;
            logFilePath = logFilePath + partiallogFilePath;
            logFileInfo = new FileInfo(logFilePath);
            logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
            if (!logDirInfo.Exists)
                logDirInfo.Create();
            if (!logFileInfo.Exists)
            {
                fileStream = logFileInfo.Create();
            }
            else
            {
                fileStream = new FileStream(logFilePath, FileMode.Append);
            }
            log = new StreamWriter(fileStream);
            log.WriteLine(logmessage);
            log.Close();
        }
    }
}
