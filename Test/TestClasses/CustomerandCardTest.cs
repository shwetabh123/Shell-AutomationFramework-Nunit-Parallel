using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Test.Base;
using Test.Config;
using OpenQA.Selenium;
using TestContext = NUnit.Framework.TestContext;
using System.Windows.Forms;
using AutoItX3Lib;
using Test.Helpers;
using Test.Pages;

namespace Test.TestClasses
{
    [TestFixture]
    [Parallelizable(ParallelScope.Fixtures)]
    public class CustomerandCardTest : BaseClass
    {
        string sub_browser = null;
        public static string customerERP = "";
        List<string> customer = new List<string>();
        public CustomerandCardTest() : base(BrowserType.Chrome)
        {
        }
        //public CustomerandCardTest(string browser,string subbrowser) : base(browser)
        //{
        //    sub_browser = subbrowser;
        //}
        //Nunit
        //[TestCase ("Fleet","Gold"),Category("Smoke"),Order(1)]
        //[TestCase("CRT","Silver"), Category("Smoke")]
        //[Test , Category("Smoke"),Order(1)]
        [TestCase(Ignore ="no code"), Category("Smoke"),Order(1)]
      //  public void CreateTopLevelCustomerDetails(string LineOfBusiness,string Band)
        public void CreateTopLevelCustomerDetails()
        {
            int srow = ConfigReader.srow;
            int erow = ConfigReader.erow;
            for (int i=srow;i<=erow;i++)
            {
                try
                {

                    //var autoIT=new AutoItX3();
                    //IE
                    System.Diagnostics.Process.Start(@"D:\Shell Framework\AutomationFramework -Nunit - Parallel\AutomationFramework\Autoitscript\HandleAuthenticationWindow.exe");
                    DriverContext.GetDriver<IWebDriver>().Navigate().GoToUrl("http://ws08r2/gfnlaunch");

                    //chrome
                  //  DriverContext.GetDriver<IWebDriver>().Navigate().GoToUrl("http://shwet:Welcome_068@ws08r2/gfnlaunch");

                    //var driverHandles = DriverContext.GetDriver<IWebDriver>().WindowHandles;
                    //var test = autoIT.Winactive;
                    //SendKeys.Send(@"code\shwet");

                 //   DriverContext.GetDriver<IWebDriver>() = GetChromeDriver();


                    DriverContext.InitDriver(_browsertype);
                    var autoIT = new AutoItX3();
                    string title = DriverContext.GetDriver<IWebDriver>().Title;

                    string testmethodname = TestContext.CurrentContext.Test.Name;

                    LogHelper.WriteLog(ConfigReader.logFilePath,"****************************");

                    LogHelper.WriteLog(ConfigReader.logFilePath,"Executing Test Case-> "+ testmethodname);

                    LogHelper.WriteLog(ConfigReader.logFilePath, "****************************");


                    Customer c = new Customer();
                    c.selectcolco(ConfigReader.GetColco);

                    string customerERP = c.CreateCustomerDetails(ConfigReader.GetColco,i);


                    DriverContext.CloseDriver();



                }
                catch(Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
            }
        }
        //MsTest
      //  [TestMethod]
        //Nunit
        //[TestCase("Fleet", "Gold","Philipines FLT 7002861"), Category("Smoke"), Order(2)]
        //[TestCase("CRT", "Silver","Philipines CRT 7077861"), Category("Smoke")]

        [Test, Category("Smoke"),Order(2)]
           //public void TopLevelCustomerEntiretyCheck(string LineOfBusiness, string Band,string cardtype)
        public void TopLevelCustomerEntiretyCheck()


        {
            int srow = ConfigReader.srow;
            int erow = ConfigReader.erow;
            for (int i = srow; i <= erow; i++)
            {
                try
                {

                    //var autoIT=new AutoItX3();
                    //IE
                    System.Diagnostics.Process.Start(@"D:\Shell Framework\AutomationFramework -Nunit - Parallel\AutomationFramework\Autoitscript\HandleAuthenticationWindow.exe");
                    DriverContext.GetDriver<IWebDriver>().Navigate().GoToUrl("http://ws08r2/gfnlaunch");

                    //chrome
                    //  DriverContext.GetDriver<IWebDriver>().Navigate().GoToUrl("http://shwet:Welcome_068@ws08r2/gfnlaunch");

                    //var driverHandles = DriverContext.GetDriver<IWebDriver>().WindowHandles;
                    //var test = autoIT.Winactive;
                    //SendKeys.Send(@"code\shwet");

                    //   DriverContext.GetDriver<IWebDriver>() = GetChromeDriver();


                    DriverContext.InitDriver(_browsertype);
                    var autoIT = new AutoItX3();
                    string title = DriverContext.GetDriver<IWebDriver>().Title;

                    string testmethodname = TestContext.CurrentContext.Test.Name;

                    LogHelper.WriteLog(ConfigReader.logFilePath, "****************************");

                    LogHelper.WriteLog(ConfigReader.logFilePath, "Executing Test Case-> " + testmethodname);

                    LogHelper.WriteLog(ConfigReader.logFilePath, "****************************");

                    Pricing c = new Pricing();

                    //PRICERULE
            


                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
            }
        }
        //MsTest
        //  [TestMethod]
        //Nunit
        [Test, Category("Smoke"),Order(3)]
        public void ChangeandVerifyCustomerStatus()
        {
            int srow = ConfigReader.srow;
            int erow = ConfigReader.erow;
            for (int i = srow; i <= erow; i++)
            {
                try
                {











                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }
            }
        }
    }
}
