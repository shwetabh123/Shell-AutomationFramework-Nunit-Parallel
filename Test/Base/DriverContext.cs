using AutoItX3Lib;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Test.Config;

namespace Test.Base
{
    public static class DriverContext
    {
        private static ThreadLocal<object> storedDriver = new ThreadLocal<object>();
        public static DriverType GetDriver<DriverType>()
        {
            return (DriverType)DriverStored;
        }
        public static object DriverStored
        {
            get
            {
                if (storedDriver == null)
                    return null;
                else return storedDriver.Value;
            }
            set
            {
                storedDriver.Value = value;
            }
        }

        public static void InitDriver(BrowserType browsertype,ConfigReader config=null)
        {


            var autoIT = new AutoItX3();

            if(browsertype == BrowserType.Chrome)
            {

                DriverStored = BaseClass.GetChromeDriver();

                DriverContext.GetDriver<IWebDriver>().Manage().Window.Maximize();


                DriverContext.GetDriver<IWebDriver>().Manage().Timeouts().PageLoad = (TimeSpan.FromSeconds(ConfigReader.GetPageLoadTimeOut()));

                DriverContext.GetDriver<IWebDriver>().Manage().Timeouts().ImplicitWait = (TimeSpan.FromSeconds(ConfigReader.GetElementLoadTimeOut()));

                DriverContext.GetDriver<IWebDriver>().Navigate().GoToUrl(ConfigReader.GetUrlChrome);

            }



            else if (browsertype == BrowserType.IExplorer)

            {
                DriverStored = BaseClass.GetIEDriver();
                DriverContext.GetDriver<IWebDriver>().Manage().Window.Maximize();

                DriverContext.GetDriver<IWebDriver>().Manage().Timeouts().PageLoad = (TimeSpan.FromSeconds(ConfigReader.GetPageLoadTimeOut()));

                DriverContext.GetDriver<IWebDriver>().Manage().Timeouts().ImplicitWait = (TimeSpan.FromSeconds(ConfigReader.GetElementLoadTimeOut()));

                System.Diagnostics.Process.Start("D://Shell Framework//AutomationFramework -Nunit - Parallel//AutomationFramework//Autoitscript//HandleAuthenticationWindow.exe");

                DriverContext.GetDriver<IWebDriver>().Navigate().GoToUrl(ConfigReader.GetUrlIE);

             //   childTest = parentTest.CreateNode(TestContext.CurrentContet.Test.MethodName);
            }


        }


        public static void CloseDriver()
        {
            IWebDriver driver=(IWebDriver)DriverStored;
            driver.Quit();
            DriverStored = null;


        }


        //public static IWebDriver Driver
        //{
        //    get
        //    {
        //        return driver;
        //    }
        //    set
        //    {
        //        driver = value;
        //    }
        //}
        //public static Browser Browser { get; set; }
    }
}
