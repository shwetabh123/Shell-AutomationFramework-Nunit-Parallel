using OpenQA.Selenium;
using Test.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Test.Helpers
{
    public class JavaScriptExecutor
    {

        public static object ExecuteScript(string script)
        {
            IJavaScriptExecutor executor = ((IJavaScriptExecutor)DriverContext.GetDriver<IWebDriver>());
           
            return executor.ExecuteScript(script);

        }

        public static object ExecuteScript(string script, params object[] args)
        {
            IJavaScriptExecutor executor = ((IJavaScriptExecutor)DriverContext.GetDriver<IWebDriver>());
            return executor.ExecuteScript(script, args);
        }

        public static void ScrollToAndClick(IWebElement element)
        {
            ExecuteScript("window.scrollTo(0," + element.Location.Y + ")");
            Thread.Sleep(300);
            element.Click();
        }

        public static void ScrollToAndClick(By locator)
        {
            IWebElement element = GenericHelper.GetElement(locator);
            ExecuteScript("window.scrollTo(0," + element.Location.Y + ")");
            Thread.Sleep(300);
            element.Click();
        }



    }
}
