using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Test.Base;
namespace Test.Pages
{
   public  class Billing
    {
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Start { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Transaction { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement BillingAcceptance { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement BillingReportDate { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement BillingPreview { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement BillingCutOff { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement BillingRefresh { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement BillingSignOf { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement BillingSignOf_Confirmation { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Click_No { get; set; }
        [FindsBy(How = How.XPath, Using = "")]
        private IWebElement Preview { get; set; }
        public Billing()
        {
            PageFactory.InitElements(DriverContext.GetDriver<IWebDriver>(), this);
        }
        public void ClickBillingPreview()
        {
            BillingPreview.Click();
        }
        public void ClickBillingCutOff()
        {
            BillingCutOff.Click();
        }
        public void ClickBillingRefreshafterBillingCutOff()
        {
           while(!(BillingSignOf.Enabled))
            {
                BillingRefresh.Click();
            }
        }
        public void ClickBillingRefreshafterBillingPreview()
        {
            while (!(BillingPreview.Enabled && BillingCutOff.Enabled))
            {
                BillingRefresh.Click();
            }
        }
        public void ClickBillingSignOf()
        {
            BillingSignOf.Click();
        }
        public void ClickBillingSignOf_Confirmation()
        {
            BillingSignOf_Confirmation.Click();
        }
        public void ClickBillingRefreshafterBillingSignOf()
        {
            while (!(Preview.Enabled))
            {
                BillingRefresh.Click();
            }
        }
        public void Click_Preview()
        {
            Preview.Click();
        }
    }
}
