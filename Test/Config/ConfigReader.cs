using System;
using System.Configuration;
using System.Xml.XPath;
using System.IO;
namespace Test.Config
{
    public class ConfigReader
    {
        public static string logFilePath = ConfigurationManager.AppSettings["logFilePath"].ToString();
        public static string downloadFilepath = ConfigurationManager.AppSettings["downloadFilepath"].ToString();
        public static string TestDataFilepath = ConfigurationManager.AppSettings["TestDataFilepath"].ToString();
        public static int srow = Int32.Parse(ConfigurationManager.AppSettings["srow"].ToString());
        public static int erow = Int32.Parse( ConfigurationManager.AppSettings["erow"].ToString());
        public static string DBServerName = ConfigurationManager.AppSettings["DBServerName"].ToString();
        public static string PH_OLTP_DB = ConfigurationManager.AppSettings["PH_OLTP_DB"].ToString();
        public static string PH_BATCH_DB = ConfigurationManager.AppSettings["PH_BATCH_DB"].ToString();
        public static string GetUrlChrome = ConfigurationManager.AppSettings["UrlChrome"].ToString();
        public static string GetUrlIE = ConfigurationManager.AppSettings["UrlIE"].ToString();
        public static string GetBrowser = ConfigurationManager.AppSettings["Browser"].ToString();
        public static string GetColco = ConfigurationManager.AppSettings["Colco"].ToString();
        public static string GetDBPassword = ConfigurationManager.AppSettings["DBPassword"].ToString();
        public static string GetDBUserName = ConfigurationManager.AppSettings["DBUserName"].ToString();
        public static int GetElementLoadTimeOut()
        {
            string timeout = ConfigurationManager.AppSettings["ElementLoadTimeout"].ToString();
            if (timeout == null)
                return 30;
            return Convert.ToInt32(timeout);
        }
        public static int GetPageLoadTimeOut()
        {
            string timeout = ConfigurationManager.AppSettings["PageLoadTimeout"].ToString();
            if (timeout == null)
                return 30;
            return Convert.ToInt32(timeout);
        }
        //public static string GetDBPassword()
        //{
        //    return ConfigurationManager.AppSettings["DBPassword"].ToString();
        //}
        //public static string GetDBUserName()
        //{
        //    return ConfigurationManager.AppSettings["DBUserName"].ToString();
        //}
        //public static string userName()
        //{
        //    return ConfigurationManager.AppSettings["userName"].ToString();
        //}
        //public static string password()
        //{
        //    return ConfigurationManager.AppSettings["password"].ToString();
        //}
        //public static void SetFrameworkSettings()
        //{
        //    XPathItem aut;
        //    XPathItem testtype;
        //    XPathItem islog;
        //    XPathItem isreport;
        //    XPathItem buildname;
        //    XPathItem logPath;
        //    string strFilename = Environment.CurrentDirectory.ToString() + "\\Config\\GlobalConfig.xml";
        //    FileStream stream = new FileStream(strFilename, FileMode.Open);
        //    XPathDocument document = new XPathDocument(stream);
        //    XPathNavigator navigator = document.CreateNavigator();
        //    //Get XML Details and pass it in XPathItem type variables
        //    aut = navigator.SelectSingleNode("EAAutoFramework/RunSettings/AUT");
        //    buildname = navigator.SelectSingleNode("EAAutoFramework/RunSettings/BuildName");
        //    testtype = navigator.SelectSingleNode("EAAutoFramework/RunSettings/TestType");
        //    islog = navigator.SelectSingleNode("EAAutoFramework/RunSettings/IsLog");
        //    isreport = navigator.SelectSingleNode("EAAutoFramework/RunSettings/IsReport");
        //    logPath = navigator.SelectSingleNode("EAAutoFramework/RunSettings/LogPath");
        //    //Set XML Details in the property to be used accross framework
        //    Settings.AUT = aut.Value.ToString();
        //    Settings.BuildName = buildname.Value.ToString();
        //    Settings.TestType = testtype.Value.ToString();
        //    Settings.IsLog = islog.Value.ToString();
        //    Settings.IsReporting = isreport.Value.ToString();
        //    Settings.LogPath = logPath.Value.ToString();
        //}
    }
}
