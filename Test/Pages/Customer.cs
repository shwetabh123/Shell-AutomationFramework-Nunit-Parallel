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
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Address { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement NewAddress { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement AddressTextBox { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Address_City { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Address_Region { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Address_PostalCode { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Address_AddressType_Main { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Address_AddressType_Registered { get; set; }

        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Address_AddressType_Correspondence { get; set; }

        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Customer_HeaderTab { get; set; }

        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Contacts { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Contacts_NewContact { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Contacts_ForeName { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Contacts_SurName { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Contacts_Telephone { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Contacts_Email { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Contacts_PinDelivery { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Contacts_CardDelivery { get; set; }
     


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
       // public string CreateCustomerDetails(string colconame, int i, string lineofbusiness, string band)
            public string CreateCustomerDetails(string colconame, int i)
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
        public void CustomerAddress(int i)
        {
            string fileName = ConfigReader.TestDataFilepath;
            ExcelHelper eat = new ExcelHelper(fileName);
            string strWorksheetName = eat.getExcelSheetName(fileName,2);
            Actions action = new Actions(DriverContext.GetDriver<IWebDriver>());
            action.MoveToElement(CustomerDetails).Build().Perform();
            action.MoveToElement(Address).Build().Perform();
            Address.Click();
            NewAddress.Click();
            TextBoxHelper.TypeInTextBox(AddressTextBox,eat.GetCellData("Address","AddressTextBox",i));
            TextBoxHelper.TypeInTextBox(Address_City, eat.GetCellData("Address", "Address_city", i));

            ComboBoxHelper.SelectElementByText(Address_Region,eat.GetCellData("Address","Address_Region",i));

            TextBoxHelper.TypeInTextBox(Address_PostalCode, eat.GetCellData("Address", "Address_PostalCode", i));

            Address_AddressType_Main.Click();

            Address_AddressType_Registered.Click();

            Address_AddressType_Correspondence.Click();

            Save.Click();

        }



        public void CustomerContacts(int i)
        {
            string fileName = ConfigReader.TestDataFilepath;
            ExcelHelper eat = new ExcelHelper(fileName);
            string strWorksheetName = eat.getExcelSheetName(fileName, 3);

            Actions action = new Actions(DriverContext.GetDriver<IWebDriver>());
            action.MoveToElement(Customer_HeaderTab).Build().Perform();
            action.MoveToElement(Contacts).Build().Perform();

            Contacts.Click();

            Contacts_NewContact.Click();
            TextBoxHelper.TypeInTextBox(Contacts_ForeName, eat.GetCellData("Contacts", "Contacts_ForeName", i));

            TextBoxHelper.TypeInTextBox(Contacts_SurName, eat.GetCellData("Contacts", "Contacts_SurName", i));

            TextBoxHelper.TypeInTextBox(Contacts_Telephone, eat.GetCellData("Contacts", "Contacts_Telephone", i));

            TextBoxHelper.TypeInTextBox(Contacts_Email, eat.GetCellData("Contacts", "Contacts_Email", i));

            Contacts_PinDelivery.Click();

            Contacts_CardDelivery.Click();

            Save.Click();






        }

    }
    }
