using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Test.Base;
using Test.Config;
using Test.Helpers;
namespace Test.Pages
{
    public class Transactions
    {
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Start { get; set; }
        public Transactions()
        {
            PageFactory.InitElements(DriverContext.GetDriver<IWebDriver>(), this);
        }
        public string CreateManualSales(string card, string expdate, string salesdate, string colconame, string customerid)
        {
            string salesitemid = "";
            return salesitemid;
        }
        public static string WritetoDX026File(string datFilePath, string recordtype, string transacionid, string transactionlinenumber, string clientcompanynumber,
            string terminalnumber, string sitenumber, string sitename, string PANEntrymethod, string saleslocation, string ismanual, string PAN,
            string cardexpirydate, string salesdatetime, string transactioncurrencycode, string siteamountperunit, string siteamountgross, string siteamountvat,
            string productcode, string subproductcode, string subproductdescription, string productstockstatus, string quantity, string vouchernumber,
            string tranproviderid, string randomnumber)
        {
            DateTime NOW = DateTime.Now;
            LogHelper.WriteLog(ConfigReader.logFilePath, "DateTime is : " + NOW);
            DateTime oDateTime = Convert.ToDateTime(NOW);
            string DateTimenew = oDateTime.ToString("yyyyMMddmmss");
            string Date = DateTimenew.Substring(0, 8);
            string Time = DateTimenew.Substring(8, 6);
            LogHelper.WriteLog(ConfigReader.logFilePath, "Date is " + Date);
            LogHelper.WriteLog(ConfigReader.logFilePath, "Time   is " + Time);
            string FileTrailer = "T|DX026_GFN_TRX_" + tranproviderid + "_" + randomnumber + "_" + Date + "_" + Time + ".dat|" + quantity + "|" + siteamountgross;
            string datFileName = "DX026_GFN_TRX_" + tranproviderid + "_" + randomnumber + "_" + Date + "_" + Time;
            string file = datFilePath + "\\" + datFileName + ".dat";
            if (!(File.Exists(file)))
            {
                using (StreamWriter sw = File.CreateText(file))
                {
                    sw.WriteLine(recordtype + "|" + transacionid + "|" + transactionlinenumber + "|" + clientcompanynumber + "|" + terminalnumber + "|" + sitenumber +
                        "|" + sitename + "|" + PANEntrymethod + "|" + saleslocation + "|" + ismanual + "|" + PAN + "|" + cardexpirydate + "|" + salesdatetime + "|" +
                        transactioncurrencycode + "|" + siteamountperunit + "|" + siteamountgross + "|" + siteamountvat + "|" + productcode + "|" + subproductcode +
                        "|" + subproductdescription + "|" + productstockstatus + "|" + quantity + "|" + vouchernumber +
                        "|||||y|additional 1|additional 2|additional 3|additional 4");
                    sw.WriteLine(FileTrailer);
                }
            }
                Thread.Sleep(246000);
                //BATCH
                DBHelper.executesqlandwritetocsvmultipletables(DriverContext.GetDriver<IWebDriver>(), ConfigReader.downloadFilepath, "sql", "BATCH",
                    "select VoucherNumber,CustomerID,SalesItemID from SalesIemUnbilled where VoucherNumber='" + vouchernumber + "' and PAN='" + PAN + "'");
                Thread.Sleep(46000);
                string salesitemid = ExcelHelper.Readcsvcolumnnamerownumberwise(DriverContext.GetDriver<IWebDriver>(), ConfigReader.downloadFilepath, "sql", "SalesItemID", 1);
                return salesitemid;
            }
        }
    }

