using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Test.Helpers;
namespace Test.Pages
{
    public class ManualFees
    {
        By Customer_HeaderTab;
        By ManualFeees;
        By NewManualFees;
        By SearchMenu;
        public ManualFees()
            {
            this.Customer_HeaderTab = By.XPath("");
            this.ManualFeees = By.XPath("");
            this.NewManualFees = By.XPath("");
            this.SearchMenu = By.XPath("");


            }
        public void NavigatetoandCreateManualFees(string quantity,string unitprice,string manaulfeetext)
        {
            try
            {
                string fileName = "";
                ExcelHelper eat = new ExcelHelper(fileName);
                string strWorksheetName = eat.getExcelSheetName(fileName ,12);
                GenericHelper.ExplicitWaitByXpath(SearchMenu,3000);
                TextBoxHelper.ClearandTypeinTextBox(SearchMenu,"Manual Fees");
                GenericHelper.ExplicitWaitByXpath(ManualFeees, 3000);
                GenericHelper.MouseHoverToElementbylocator(ManualFeees);
                GenericHelper.ExplicitWaitByXpath(ManualFeees, 3000);
                GenericHelper.MouseHoverToElementandClickbylocator(ManualFeees);
                GenericHelper.ExplicitWaitByXpath(NewManualFees, 3000);
                GenericHelper.MouseHoverToElementandClickbylocator(NewManualFees);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
