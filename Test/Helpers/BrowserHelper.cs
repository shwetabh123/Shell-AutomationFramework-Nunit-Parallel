using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Test.Base;

namespace Test.Helpers
{
    public  class BrowserHelper
    {

        public static void BrowserMaximize()
        {

            DriverContext.GetDriver<IWebDriver>().Manage().Window.Maximize();
        }

        public static void GoBack()
        {

            DriverContext.GetDriver<IWebDriver>().Navigate().Back();
        }

        public static void Forward()
        {

            DriverContext.GetDriver<IWebDriver>().Navigate().Forward();
        }


        public static void RefreshPage()
        {

            DriverContext.GetDriver<IWebDriver>().Navigate().Refresh();
        }
    }
}
