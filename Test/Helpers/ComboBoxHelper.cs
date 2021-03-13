using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using Test.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace Test.Helpers
{
    public class ComboBoxHelper
    {
        private static SelectElement select;
        private static WebDriverWait GetWebDriverWait(IWebDriver driver, TimeSpan timeout)
        {
            WebDriverWait wait = new WebDriverWait(driver, timeout)
            {
                PollingInterval = TimeSpan.FromMilliseconds(250)
            };
            wait.IgnoreExceptionTypes(typeof(NoSuchElementException));
            return wait;
        }
        public static void SelectElementByIndexWitWait(By locator, int index)
        {
            WebDriverWait wait = GetWebDriverWait(DriverContext.GetDriver<IWebDriver>(), TimeSpan.FromSeconds(60));
            IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(locator));
            select = new SelectElement(element);
            select.SelectByIndex(index);
        }
        public static void SelectElementByIndex(By locator, int index)
        {
            select = new SelectElement(GenericHelper.GetElement(locator));
            select.SelectByIndex(index);
        }
        //public static void SelectElementByText(By locator, string visibletext)
        //{
        //    select = new SelectElement(GenericHelper.GetElement(locator));
        //    select.SelectByText(visibletext);
        //}
        //public static void SelectElementByText(By locator, string visibletext)
        //{
        //    select = new SelectElement(DriverContext.GetDriver<IWebDriver>().FindElement(locator));
        //    select.SelectByText(visibletext);
        //}
        public static void SelectElementByText(IWebElement element, string visibletext)
        {
            select = new SelectElement(element);
            select.SelectByText(visibletext);
        }
        public static void SelectElementByText(By locator, string visibletext)
        {
            var dropdown = new SelectElement(DriverContext.GetDriver<IWebDriver>().FindElement(locator));
            dropdown.SelectByText(visibletext);
        }
        public static string SelectInDropDownByText(IWebElement element, string visibletext)
        {
            var dropdown = new SelectElement(element);
            dropdown.SelectByText(visibletext);
            return visibletext;
        }
        //public static void SelectElementByValue(By locator, string valueTexts)
        //{
        //    select = new SelectElement(GenericHelper.GetElement(locator));
        //    select.SelectByValue(valueTexts);
        //}
        public static void SelectElementByValue(By locator, string valueTexts)
        {
            select = new SelectElement(DriverContext.GetDriver<IWebDriver>().FindElement(locator));
            select.SelectByValue(valueTexts);
        }
        public static void SelectElementByValue(IWebElement element, string visibletext)
        {
            select = new SelectElement(element);
            select.SelectByValue(visibletext);
        }
        public static IList<string> GetAllItem(By locator)
        {
            select = new SelectElement(GenericHelper.GetElement(locator));
            return select.Options.Select((x) => x.Text).ToList();
        }
    }
}
