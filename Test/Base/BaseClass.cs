using AutoItX3Lib;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports.Reporter.Configuration;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Test.Config;


namespace Test.Base
{
    [TestFixture]
    public class BaseClass
    {
        public static ExtentReports extent;
        public static ExtentTest test;
        public static ExtentHtmlReporter htmlReporter;
        public static ExtentTest parentTest;
        public static ExtentTest childTest;
        public BrowserType _browsertype;

        //***********************************//

        //For only MSTEST  Framework
        //private TestContext testContextInstance;
        //public TestContext TestContext
        //{

        //    get { return testContextInstance; }

        //    set { testContextInstance = value; }

        //}
            


        //****************************************//

        
        //using enumerator
        public BaseClass(BrowserType browser)
        {
            _browsertype = browser;
        }
        public BaseClass(string browser)
        {
            browser = ConfigReader.GetBrowser;
        }
        public static FirefoxProfile GetFirefoxptions()
        {
            FirefoxProfile profile = new FirefoxProfile();
            FirefoxProfileManager manager = new FirefoxProfileManager();
            profile = manager.GetProfile("default");
            //  Logger.Info(" Using Firefox Profile ");
            return profile;
        }
        private static ChromeOptions GetChromeOptions()
        {
            ChromeOptions option = new ChromeOptions();
            option.AddArgument("start-maximized");
            option.AddArgument("--headless");
            option.AddExtension(@"C:\Users\rahul.rathore\Desktop\Cucumber\extension_3_0_12.crx");
            //  Logger.Info(" Using Chrome Options ");
            return option;
        }
        private static InternetExplorerOptions GetIEOptions()
        {
            InternetExplorerOptions options = new InternetExplorerOptions();
            options.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
            options.EnsureCleanSession = true;
            options.EnablePersistentHover = true;
            options.EnableNativeEvents = true;
            options.RequireWindowFocus = true;
            options.ElementScrollBehavior = InternetExplorerElementScrollBehavior.Bottom;
            //   Logger.Info(" Using Internet Explorer Options ");
            return options;
        }
        //private static FirefoxDriver GetFirefoxDriver()
        //{
        //    FirefoxOptions options = new FirefoxOptions();
        //    FirefoxDriver driver = new FirefoxDriver(GetFirefoxptions());
        //    return driver;
        //}
        public static ChromeDriver GetChromeDriver()
        {
            ChromeDriver driver = new ChromeDriver(GetChromeOptions());
            return driver;
        }
        public static InternetExplorerDriver GetIEDriver()
        {
            InternetExplorerDriver driver = new InternetExplorerDriver(GetIEOptions());
            return driver;
        }
        //private static PhantomJSDriver GetPhantomJsDriver()
        //{
        //    PhantomJSDriver driver = new PhantomJSDriver(GetPhantomJsDrvierService());
        //    return driver;
        //}
        //private static PhantomJSOptions GetPhantomJsptions()
        //{
        //    PhantomJSOptions option = new PhantomJSOptions();
        //    option.AddAdditionalCapability("handlesAlerts", true);
        //    //     Logger.Info(" Using PhantomJS Options  ");
        //    return option;
        //}
        //private static PhantomJSDriverService GetPhantomJsDrvierService()
        //{
        //    PhantomJSDriverService service = PhantomJSDriverService.CreateDefaultService();
        //    service.LogFile = "TestPhantomJS.log";
        //    service.HideCommandPromptWindow = false;
        //    service.LoadImages = true;
        //    //   Logger.Info(" Using PhantomJS Driver Service  ");
        //    return service;
        //}
        public static string Capture(string ScreenShotName)
        {
            DateTime dt = DateTime.Now; // Or whatever
            string dateName = dt.ToString("yyyyMMddHHmmss");
            ITakesScreenshot ts = (ITakesScreenshot)DriverContext.GetDriver<IWebDriver>();
            Screenshot Screenshot = ts.GetScreenshot();
            string path = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            string uptobinpath = path.Substring(0, path.LastIndexOf("bin")) + "Screenshots\\" + ScreenShotName + dateName + ".png";
            //    +DateTime.Now.ToString(“Dd_MMMM_hh_mm_ss_tt”) + “.Png”;
            string localpath = new Uri(uptobinpath).LocalPath;
            Screenshot.SaveAsFile(localpath, ScreenshotImageFormat.Png);
            return localpath;
        }
        //MSTEST
        //    [AssemblyInitialize]
        //      public static void SetupTests(TestContext TestContext)
        //NUNIT
        [OneTimeSetUp]
        public void Setup()
        {
            // Relevantcodes extent Report 2.41
            //*******************************************************************************************
            // //To obtain the current solution path/project path
            // string pth = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            // string actualPath = pth.Substring(0, pth.LastIndexOf("bin"));
            // string projectPath = new Uri(actualPath).LocalPath;
            // Console.WriteLine(projectPath);
            // //Append the html report file to current project path
            // string reportPath = projectPath + "Report\\TestRunReport.html";
            // Console.WriteLine(reportPath);
            // //Boolean value for replacing exisisting report
            //extent = new ExtentReports(reportPath, true);
            // //Add QA system info to html report
            // extent.AddSystemInfo("Host Name", "Shwetabh")
            //     .AddSystemInfo("Environment", "Stage")
            //    .AddSystemInfo("Username", "shwetabh123");
            // //Adding config.xml file
            // extent.LoadConfig(projectPath + "extent-config.xml"); //Get the config.xml file from http://extentreports.com
            //******************************************
            //aventstack extentreport 3.0.0
            //****************************************************
            //To obtain the current solution path/project path
            //string pth = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
            // string pth = System.Environment.CurrentDirectory;
            //  string pth = AppDomain.CurrentDomain.BaseDirectory;
            //-----------------------
            //string pth = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            //string actualPath = pth.Substring(0, pth.LastIndexOf("bin"));
            //string projectPath = new Uri(actualPath).LocalPath;
            //Console.WriteLine(projectPath);

            //-------------------
            string projectPath = (@"D:\Shell Framework");

            //Append the html report file to current project path
            string reportPath = projectPath + "Report\\TestRunReport.html";
            //string reportPath = "TestRunReport.html";
            Console.WriteLine(reportPath);
            //aventstack
            //  ExtentHtmlReporter htmlReporter = new ExtentHtmlReporter(reportPath);
            //Boolean value for replacing exisisting report
            //relevent codes
           // extent = new ExtentReports();
            //aventstack
            //   extent = new ExtentReports();
            //// report title
              htmlReporter.Configuration().DocumentTitle = "aventstack - ExtentReports";
            //// encoding, default = UTF-8
              htmlReporter.Configuration().Encoding = "UTF-8";
              htmlReporter.Configuration().Theme = AventStack.ExtentReports.Reporter.Configuration.Theme.Standard;
            //// report or build name
            //   htmlReporter.Config.ReportName = "Build-1224";
            //// chart location - top, bottom
            htmlReporter.Configuration().ChartLocation = ChartLocation.Top;
            //// theme - standard, dark
            ////htmlReporter.Configuration().Theme = Theme.Dark;
            //// add custom css
            htmlReporter.Configuration().CSS = "css-string";
            //// add custom javascript
            htmlReporter.Configuration().JS = "js-string";
            //// create ExtentReports and attach reporter(s)
           
            //aventstack
            extent = new ExtentReports();
            extent.AttachReporter(htmlReporter);
            extent.AddSystemInfo("Platform", "Windows");
            extent.AddSystemInfo("Host Name", "localhost");
            extent.AddSystemInfo("Environment", "QA");
            extent.AddSystemInfo("User Name", "testUser");
            //************
            //Gives full path package---class name---test name
            // parentTest = extent.CreateTest(TestContext.CurrentContext.Test.ClassName);
            //Gives only class name----use this
            //aventstack
            parentTest = extent.CreateTest(TestContext.CurrentContext.Test.Name);
            childTest = parentTest.CreateNode(TestContext.CurrentContext.Test.Name);
          //  parenttest = extent.StartTest("Parent", "Test Started");
        }


      //  MSTEST
      //     [TestInitialize]
       // Nunit
        [SetUp]
        public void BeforeTest()
        {
           // DriverContext.InitDriver(_browsertype);
        }


        ////MSTEST
        ////   [TestInitialize]
        ////Nunit
        //[SetUp]
        //public void BeforeTest()
        //{
        //    OpenBrowser(_browsertype);
        //}
        //public void OpenBrowser(BrowserType browsertype)
        //{
        //    if (browsertype == BrowserType.Chrome)
        //    {
        //        DriverContext.Driver = GetChromeDriver();
        //        DriverContext.Driver.Manage().Window.Maximize();
        //        DriverContext.Driver.Manage().Timeouts().PageLoad = (TimeSpan.FromSeconds(ConfigReader.GetPageLoadTimeOut()));
        //        DriverContext.Driver.Manage().Timeouts().ImplicitWait = (TimeSpan.FromSeconds(ConfigReader.GetElementLoadTimeOut()));
        //        var autoIT = new AutoItX3();
        //        DriverContext.Driver.Navigate().GoToUrl(ConfigReader.GetUrlChrome);
        //        childTest = parentTest.CreateNode(TestContext.CurrentContext.Test.MethodName);
        //    }
        //    else if (browsertype == BrowserType.IExplorer)
        //    {
        //        DriverContext.Driver = GetChromeDriver();
        //        DriverContext.Driver.Manage().Window.Maximize();
        //        DriverContext.Driver.Manage().Timeouts().PageLoad = (TimeSpan.FromSeconds(ConfigReader.GetPageLoadTimeOut()));
        //        DriverContext.Driver.Manage().Timeouts().ImplicitWait = (TimeSpan.FromSeconds(ConfigReader.GetElementLoadTimeOut()));
        //        var autoIT = new AutoItX3();
        //        System.Diagnostics.Process.Start(@"D:\Shell Framework\AutomationFramework -Nunit\AutomationFramework\Autoitscript\HandleAuthenticationWindow.exe");
        //        DriverContext.Driver.Navigate().GoToUrl(ConfigReader.GetUrlIE);
        //        childTest = parentTest.CreateNode(TestContext.CurrentContext.Test.MethodName);
        //    }
        //}


        ////  MSTest
        //[TestCleanup]
        //////NUNIT
        ////[TearDown]
        //public static void TestCleanup(TestContext TestContext)
        //{
        //    var status = TestContext.CurrentTestOutcome;
        //    if (status == UnitTestOutcome.Failed)
        //    {
        //        string screenShotPath = BaseClass.Capture( "screesnshotname");
        //        childtest.Log(LogStatus.Pass, "Test Failed");
        //        childtest.Log(LogStatus.Fail, "Snapshot below: " + childtest.AddScreenCapture(screenShotPath));
        //    }
        //    else if (status == UnitTestOutcome.Passed)
        //    {
        //        string screenShotPath = BaseClass.Capture("screesnshotname");
        //        childtest.Log(LogStatus.Pass, "Test Passed");
        //        childtest.Log(LogStatus.Pass, "Snapshot below: " + childtest.AddScreenCapture(screenShotPath));
        //    }
        //    else if (status == UnitTestOutcome.Error)
        //    {
        //        string screenShotPath = BaseClass.Capture("screesnshotname");
        //        childtest.Log(LogStatus.Pass, "Test Failed");
        //        childtest.Log(LogStatus.Error, "Snapshot below: " + childtest.AddScreenCapture(screenShotPath));
        //    }
        //    else if (status == UnitTestOutcome.Unknown)
        //    {
        //        string screenShotPath = BaseClass.Capture("screesnshotname");
        //        childtest.Log(LogStatus.Pass, "Test Skipped");
        //        childtest.Log(LogStatus.Skip, "Snapshot below: " + childtest.AddScreenCapture(screenShotPath));
        //    }
        //    if (DriverContext.Driver != null)
        //    {
        //        //End test report
        //        extent.EndTest(childtest);
        //        DriverContext.Driver.Close();
        //        DriverContext.Driver.Quit();
        //    }
        //    //    Logger.Info(" Stopping the Driver  ");
        //}
        //MSTest
        // [ClassCleanup]
        //  [AssemblyCleanup]
        //NUnit
        [OneTimeTearDown]
        public static void TearDown()
        {
            BaseClass.extent.Flush();
        }
    }
}
