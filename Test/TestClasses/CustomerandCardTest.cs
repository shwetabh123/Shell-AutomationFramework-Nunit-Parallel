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

        [TestCase ("Fleet","Gold"),Category("Smoke"),Order(1)]

        [TestCase("CRT","Silver"), Category("Smoke")]
        public void CreateTopLevelCustomerDetails(string LineOfBusiness,string Band)
        {
            int srow = ConfigReader.srow;

            int erow = ConfigReader.erow;

            for (int i=srow;i<=erow;i++)
            {

                try
                {

                    DriverContext.InitDriver(_browsertype);


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

        [TestCase("Fleet", "Gold","Philipines FLT 7002861"), Category("Smoke"), Order(2)]

        [TestCase("CRT", "Silver","Philipines CRT 7077861"), Category("Smoke")]
        public void TopLevelCustomerEntiretyCheck(string LineOfBusiness, string Band,string cardtype)
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
