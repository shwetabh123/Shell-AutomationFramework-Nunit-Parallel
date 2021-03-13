using Test.Base;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
namespace Test.Helpers
{
    public class TextBoxHelper
    {
        private static IWebElement element;
        public static void TypeInTextBox(By locator, string text)
        {
            element = GenericHelper.GetElement(locator);
            element.SendKeys(text);
        }
        public static void TypeInTextBox(IWebElement element, string text)
        {
            element.Clear();
            element.SendKeys(text);
        }
        public static void TypeInTextBoxusingJavaScript(string id, string text)
        {
         IJavaScriptExecutor executor =(IJavaScriptExecutor)DriverContext.GetDriver<IWebDriver>();
            executor.ExecuteScript("document.getElementById(id).value='text'");
        }
        public static void ClearTextBox(By locator)
        {
            element = GenericHelper.GetElement(locator);
            element.Clear();
        }
        public static void ClearTextBox(IWebElement element)
        {
            element.Clear();
        }
        public static void ClearandTypeinTextBox(IWebElement element, string text)
        {
            element.Clear();
            Thread.Sleep(3000);
            element.SendKeys(text);
        }
        public static void ClearandTypeinTextBox(By locator, string text)
        {
            DriverContext.GetDriver<IWebDriver>().FindElement(locator).Clear();
            Thread.Sleep(3000);
            DriverContext.GetDriver<IWebDriver>().FindElement(locator).SendKeys(text);
        }
    }
}
