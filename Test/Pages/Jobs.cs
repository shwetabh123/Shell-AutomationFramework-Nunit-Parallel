using OpenQA.Selenium;
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
   public class Jobs
    {
    public Jobs()
        {
            PageFactory.InitElements(DriverContext.GetDriver<IWebDriver>(), this);
        }
        public void Job19forSalesItem(string customerid,string salesitemid)
        {
            string BilledSalesItemid = "";
            string JOBSTATUSID = "";
            try
            {
                //BATCH
                DBHelper.executesqlandwritetocsvmultipletables(DriverContext.GetDriver<IWebDriver>(),ConfigReader.downloadFilepath,"sql","BATCH",
                    "select VoucherNumber,CustomerID,SalesItemID from SalesIemUnbilled where customerid='"+customerid+"' and salesitemid='"+salesitemid+"'");
                LogHelper.WriteLog(ConfigReader.logFilePath,"SalesItem to Bill" + salesitemid);
                //BATCH
                DBHelper.executesqlandwritetocsvmultipletables(DriverContext.GetDriver<IWebDriver>(),ConfigReader.downloadFilepath,"sql","BATCH","select Top 1 * from job where jobtypeid=19 order by 1 desc");
                string ID = ExcelHelper.Readcsvcolumnnamerownumberwise(DriverContext.GetDriver<IWebDriver>(),ConfigReader.downloadFilepath,"sql","ID",1);
                DBHelper.executesqlandwritetocsvmultipletables(DriverContext.GetDriver<IWebDriver>(),ConfigReader.downloadFilepath,"sql","BATCH",
                    "update job set NextRunDate=null where id='"+ID+"' and JobTypeID=19");
                //BATCH
                DBHelper.executesqlandwritetocsvmultipletables(DriverContext.GetDriver<IWebDriver>(), ConfigReader.downloadFilepath, "sql", "BATCH",
                    "select StatusID,ID from job where jobtypeid =19 and ID='"+ID+"'");
                JOBSTATUSID = ExcelHelper.Readcsvcolumnnamerownumberwise(DriverContext.GetDriver<IWebDriver>(), ConfigReader.downloadFilepath,"sql","StatusID",1);
                //BATCH
                DBHelper.executesqlandwritetocsvmultipletables(DriverContext.GetDriver<IWebDriver>(), ConfigReader.downloadFilepath, "sql", "BATCH",
                  "Select VoucherNumber,CustomerID,SalesItemID from SalesItem where customerid='"+customerid+"' and salesitemid='"+salesitemid+"'");
                BilledSalesItemid = ExcelHelper.Readcsvcolumnnamerownumberwise(DriverContext.GetDriver<IWebDriver>(), ConfigReader.downloadFilepath, "sql", "SalesItemID", 1);
                LogHelper.WriteLog(ConfigReader.logFilePath,"BilledSalesItemid is :-" + BilledSalesItemid);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
