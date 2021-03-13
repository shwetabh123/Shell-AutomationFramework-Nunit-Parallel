using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Test.Base;
using Test.Config;
using Test.Helpers;
namespace Test.Pages
{
    public class Customer
    {
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement selectColco { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Start { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Customers { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement CreatetopLevelCustomer { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement LineOfBusiness { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement FullName { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement ShortName { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement TradingName { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement LegalEntity { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement IncorporationDate { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement VATRegistrationNumber { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Band { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement MarketingSegmentation { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement CreditLimit { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement RequestedLimit { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Save { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement CustomerDetails { get; set; }
        public Customer()
        {
            PageFactory.InitElements(DriverContext.GetDriver<IWebDriver>(), this);
        }
        public void selectcolco(string colconame)
        {
            string xpath1 = null;
            string xpath2 = null;
            string xpath3 = null;
            string xpathgenerated = null;
            xpath1 = "";
            xpath2 = colconame;
            xpath3 = "";
            xpathgenerated = xpath1 + xpath2 + xpath3;
            IWebElement selectcolco = GenericHelper.FindElementWithXpath(xpathgenerated);
            GenericHelper.JavaScriptClick(selectcolco);
        }
        public string CreateCustomerDetails(string colconame, int i, string lineofbusiness, string band)
        {
            string customerERP = "";
            string pth = AppDomain.CurrentDomain.BaseDirectory;
            Actions action = new Actions(DriverContext.GetDriver<IWebDriver>());
            IReadOnlyCollection<string> all_windowHandles = DriverContext.GetDriver<IWebDriver>().WindowHandles;
            foreach (string handle in all_windowHandles)
            {
                DriverContext.GetDriver<IWebDriver>().SwitchTo().Window(handle);
                if (DriverContext.GetDriver<IWebDriver>().Title.Contains(colconame))
                {
                    string fileName = ConfigReader.TestDataFilepath;
                    ExcelHelper eat = new ExcelHelper(fileName);
                    string strWorksheetName = eat.getExcelSheetName(fileName, 1);
                    BrowserHelper.BrowserMaximize();
                    action.MoveToElement(Start).Build().Perform();
                    action.MoveToElement(Customers).Build().Perform();
                    action.MoveToElement(CreatetopLevelCustomer).Click().Build().Perform();
                    ComboBoxHelper.SelectInDropDownByText(LineOfBusiness, eat.GetCellData(strWorksheetName, "LineOfBusiness", i));
                    ComboBoxHelper.SelectInDropDownByText(LineOfBusiness, lineofbusiness);
                    string randomnumber = RandomNumberandStringGenerator.randomnumberwithoneargument(6);
                    string fullname = "Automation" + randomnumber;
                    TextBoxHelper.ClearandTypeinTextBox(FullName, fullname);
                    TextBoxHelper.ClearandTypeinTextBox(ShortName, fullname);
                    TextBoxHelper.ClearandTypeinTextBox(TradingName, fullname);
                    ComboBoxHelper.SelectInDropDownByText(LegalEntity, eat.GetCellData(strWorksheetName, "LegalEntity", i));
                    TextBoxHelper.ClearandTypeinTextBox(IncorporationDate, eat.GetCellData(strWorksheetName, "IncorporationDate", i));
                    IncorporationDate.SendKeys(Keys.Tab);
                    TextBoxHelper.ClearandTypeinTextBox(VATRegistrationNumber, randomnumber);
                    ComboBoxHelper.SelectInDropDownByText(Band, eat.GetCellData(strWorksheetName, "Band", i));
                    ComboBoxHelper.SelectInDropDownByText(MarketingSegmentation, eat.GetCellData(strWorksheetName, "MarketingSegmentation", i));
                    TextBoxHelper.ClearandTypeinTextBox(CreditLimit, eat.GetCellData(strWorksheetName, "CreditLimit", i));
                    TextBoxHelper.ClearandTypeinTextBox(RequestedLimit, eat.GetCellData(strWorksheetName, "RequestedLimit", i));
                    Save.Click();
                    customerERP = DriverContext.GetDriver<IWebDriver>().FindElement(By.XPath("")).GetAttribute("value");
                    LogHelper.WriteLog(ConfigReader.logFilePath, "Customer ERP is :-" + customerERP);

                    string customerERPValue = eat.SetCellData(strWorksheetName,"CustomerERP",i, customerERP);
                }
            }
            return customerERP;
        }
        public void CustomerAddress()
        {
            string fileName = ConfigReader.TestDataFilepath;
            ExcelHelper eat = new ExcelHelper(fileName);
            Actions action = new Actions(DriverContext.GetDriver<IWebDriver>());
            action.MoveToElement(CustomerDetails).Build().Perform();
        }
    }
}
