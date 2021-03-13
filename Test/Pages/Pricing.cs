using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Test.Base;
using Test.Helpers;
namespace Test.Pages
{
    public class Pricing
    {
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Start { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement PricingandFees { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement SearchPriceRule { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement NewPriceRule { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement PriceRule_Description { get; set; }
        public Pricing()
        {
            PageFactory.InitElements(DriverContext.GetDriver<IWebDriver>(), this);
        }
        public void NavigatetoSearchPriceRule(string colconame)
        {
            Actions action = new Actions(DriverContext.GetDriver<IWebDriver>());
            IReadOnlyCollection<string> all_windowHandles = DriverContext.GetDriver<IWebDriver>().WindowHandles;
            foreach(string handle in all_windowHandles)
            {
                DriverContext.GetDriver<IWebDriver>().SwitchTo().Window(handle);
                if(DriverContext.GetDriver<IWebDriver>().Title.Contains(colconame))
                    {
                    string fileName = @"";
                    ExcelHelper eat = new ExcelHelper(fileName);
                    string strWorksheetName = eat.getExcelSheetName(fileName,11);
                    int rowcountofexcel = eat.GetRowCount(strWorksheetName);
                    BrowserHelper.BrowserMaximize();
                    action.MoveToElement(Start).Build().Perform();
                    Thread.Sleep(2000);
                    action.MoveToElement(PricingandFees).Build().Perform();
                    Thread.Sleep(2000);
                    action.MoveToElement(SearchPriceRule).Click().Build().Perform();
                    Thread.Sleep(2000);
                }
            }
        }
        public void Click_NewPriceRule()
        {
            string fileName = @"";
            ExcelHelper eat = new ExcelHelper(fileName);
            string strWorksheetName = eat.getExcelSheetName(fileName, 11);
            int rowcountofexcel = eat.GetRowCount(strWorksheetName);
            GenericHelper.JavaScriptClick(NewPriceRule);
        }
    }
}
