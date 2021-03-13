using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.UI;
using Test.Base;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Workbook = Aspose.Cells.Workbook;
using Worksheet = Aspose.Cells.Worksheet;
using System.Diagnostics;
using Test.Config;
namespace Test.Helpers
{
    public class GenericHelper
    {
        public static SelectElement select;
        public static string clipboardText;
        public static string latestfile = "";
        public static string downloadFilepath;
        public static string logFilePath;
        public static string folder = null;
        public static string filepath;
        public static string partiallogFilePath = "\\" + "Log-" + System.DateTime.Now.ToString("MM-dd-yyyy_HHmmss") +
                                                  "." + "txt";
        public static bool status = false;
        private static string existingWindowHandle;
        private static Func<IWebDriver, IList<IWebElement>> GetAllElements(By locator)
        {
            return ((x) =>
            {
                return x.FindElements(locator);
            });
        }
        public static bool IsElemetPresent(By locator)
        {
            try
            {
                return DriverContext.GetDriver<IWebDriver>().FindElements(locator).Count == 1;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public static IWebElement GetElement(By locator)
        {
            if (IsElemetPresent(locator))
                return DriverContext.GetDriver<IWebDriver>().FindElement(locator);
            else
                throw new NoSuchElementException("Element Not Found : " + locator.ToString());
        }
        public static void TakeScreenShot(string filename = "Screen")
        {
            var screen = DriverContext.GetDriver<IWebDriver>().TakeScreenshot();
            if (filename.Equals("Screen"))
            {
                filename = filename + DateTime.UtcNow.ToString("yyyy-MM-dd-mm-ss") + ".jpeg";
                screen.SaveAsFile(filename, ScreenshotImageFormat.Jpeg);
                return;
            }
            screen.SaveAsFile(filename, ScreenshotImageFormat.Jpeg);
        }
        public static void Click(IWebDriver driver, string selector)
        {
            driver.FindElement(By.CssSelector(selector)).Click();
        }
        public static void Clickwithcssselector(IWebDriver driver, string selector)
        {
            driver.FindElement(By.CssSelector(selector)).Click();
        }
        public static void ClickwithID(IWebDriver driver, string ID)
        {
            driver.FindElement(By.Id(ID)).Click();
        }
        public static void ClickwithXapth(IWebDriver driver, string Xpath)
        {
            driver.FindElement(By.XPath(String.Format(Xpath))).Click();
        }
        public static void ClickwithXapth(IWebElement element)
        {
            element.Click();
        }
        public static void JavaScriptClick(IWebElement element)
        {
            IJavaScriptExecutor executor =(IJavaScriptExecutor)DriverContext.GetDriver<IWebDriver>();
            executor.ExecuteScript("argument[0].click();",element);
        }
        public static void Click(By locator)
        {
            DriverContext.GetDriver<IWebDriver>().FindElement(locator).Click();
        }
        public static void FindelementWithXpath(IWebDriver driver, string xapth)
        {
            driver.FindElement(By.XPath(xapth));
        }
        public static IWebElement FindelementWithXpath( string xapth)
        {
          return  DriverContext.GetDriver<IWebDriver>().FindElement(By.XPath(xapth));
        }
        //  ************with WAIT **********************************//
        //**************************WITH WAIT*************************//
        public static IWebElement GetElementwithWait(IWebDriver driver, string xpath, int wait)
        {
            return (new WebDriverWait(driver, TimeSpan.FromSeconds(wait))).Until(ExpectedConditions.ElementIsVisible(By.XPath(xpath)));
        }
        public static void ClickwithWait(IWebDriver driver, string selector, int wait)
        {
            new WebDriverWait(driver, TimeSpan.FromSeconds(wait)).
                Until(ExpectedConditions.ElementIsVisible(By.CssSelector(selector))).Click();
        }
        public static IWebElement GetElementswithXpath(IWebDriver driver, string xapth, int wait)
        {
            return (new WebDriverWait(driver, TimeSpan.FromSeconds(wait)))
               .Until(ExpectedConditions.ElementIsVisible(By.XPath(xapth)));
        }
        public static IWebElement GetElementwithXpath(IWebDriver driver, string xapth, int wait)
        {
            return (new WebDriverWait(driver, TimeSpan.FromSeconds(wait)))
               .Until(ExpectedConditions.ElementIsVisible(By.XPath(xapth)));
        }
        public static void ClickWithWait(IWebDriver driver, string selector, int wait)
        {
            new WebDriverWait(driver, TimeSpan.FromSeconds(wait))
                .Until(ExpectedConditions.ElementIsVisible(By.CssSelector(selector))).Click();
        }
        public static void SelectElementByIndexWitWait(By locator, int index)
        {
            WebDriverWait wait = GetWebDriverWait(DriverContext.GetDriver<IWebDriver>(), TimeSpan.FromSeconds(60));
            IWebElement element = wait.Until(ExpectedConditions.ElementIsVisible(locator));
            select = new SelectElement(element);
            select.SelectByIndex(index);
        }
        private static WebDriverWait GetWebDriverWait(IWebDriver driver, TimeSpan timeout)
        {
            WebDriverWait wait = new WebDriverWait(driver, timeout)
            {
                PollingInterval = TimeSpan.FromMilliseconds(250)
            };
            wait.IgnoreExceptionTypes(typeof(NoSuchElementException));
            return wait;
        }
        public static void WaitforElement(IWebDriver driver, string xapth, int wait)
        {
            WebDriverWait w = new WebDriverWait(driver, TimeSpan.FromSeconds(wait));
            IWebElement element = w.Until(ExpectedConditions.ElementIsVisible(By.XPath(xapth)));
        }
        //use these 
        public static void WaitforElement(By locator, int wait)
        {
            WebDriverWait w = new WebDriverWait(DriverContext.GetDriver<IWebDriver>(), TimeSpan.FromSeconds(wait));
            IWebElement element = w.Until(ExpectedConditions.ElementIsVisible(locator));
        }
        public static void WaitTime()
        {
            WebDriverWait webDriverWait = new WebDriverWait(DriverContext.GetDriver<IWebDriver>(), TimeSpan.FromSeconds(60));
        }
        public static void ExplicitWaitforElement(By locator, int wait)
        {
            try
            {
                WebDriverWait explicitwait = new WebDriverWait(DriverContext.GetDriver<IWebDriver>(), TimeSpan.FromSeconds(wait));
                explicitwait.PollingInterval = TimeSpan.FromMilliseconds(500);
                explicitwait.IgnoreExceptionTypes(typeof(NoSuchElementException));
                IWebElement element = explicitwait.Until(ExpectedConditions.ElementExists(locator));
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }
        public static void FluentWaitforElement(By locator, int wait)
        {
            try
            {
                DefaultWait<IWebDriver> fluentwait = new DefaultWait<IWebDriver>(DriverContext.GetDriver<IWebDriver>());
                fluentwait.Timeout = TimeSpan.FromSeconds(15);
                fluentwait.PollingInterval = TimeSpan.FromMilliseconds(500);
                fluentwait.IgnoreExceptionTypes(typeof(NoSuchElementException));
                IWebElement element = fluentwait.Until(ExpectedConditions.ElementExists(locator));
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }
        public static void ExplicitWaitByName(string Namelocator, int wait)
        {
            try
            {
                WebDriverWait explicitwait = new WebDriverWait(DriverContext.GetDriver<IWebDriver>(), TimeSpan.FromSeconds(wait));
                explicitwait.PollingInterval = TimeSpan.FromMilliseconds(500);
                explicitwait.IgnoreExceptionTypes(typeof(NoSuchElementException));
                IWebElement element = explicitwait.Until(ExpectedConditions.ElementExists(By.Name(Namelocator)));
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }
        public static void ExplicitWaitByXpath(string Xpathlocator, int wait)
        {
            try
            {
                WebDriverWait explicitwait = new WebDriverWait(DriverContext.GetDriver<IWebDriver>(), TimeSpan.FromSeconds(wait));
                explicitwait.PollingInterval = TimeSpan.FromMilliseconds(500);
                explicitwait.IgnoreExceptionTypes(typeof(NoSuchElementException));
                IWebElement element = explicitwait.Until(ExpectedConditions.ElementExists(By.XPath(Xpathlocator)));
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }
        public static void ExplicitWaitByXpath(By Xpathlocator, int wait)
        {
            try
            {
                WebDriverWait explicitwait = new WebDriverWait(DriverContext.GetDriver<IWebDriver>(), TimeSpan.FromSeconds(wait));
                explicitwait.PollingInterval = TimeSpan.FromMilliseconds(500);
                explicitwait.IgnoreExceptionTypes(typeof(NoSuchElementException));
                IWebElement element = explicitwait.Until(ExpectedConditions.ElementExists(Xpathlocator));
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }
        //public static void ExplicitWaitByXpath(IWebElement element, int wait)
        //{
        //    try
        //    {
        //        WebDriverWait explicitwait = new WebDriverWait(DriverContext.GetDriver<IWebDriver>(), TimeSpan.FromSeconds(wait));
        //        explicitwait.Until(ExpectedConditions.ElementExists(element));
        //    }
        //    catch (Exception e)
        //    {
        //        Debug.WriteLine(e.Message);
        //    }
        //}
        public static void ExplicitWaitById(string Idlocator, int wait)
        {
            try
            {
                WebDriverWait explicitwait = new WebDriverWait(DriverContext.GetDriver<IWebDriver>(), TimeSpan.FromSeconds(wait));
                explicitwait.PollingInterval = TimeSpan.FromMilliseconds(500);
                explicitwait.IgnoreExceptionTypes(typeof(NoSuchElementException));
                IWebElement element = explicitwait.Until(ExpectedConditions.ElementExists(By.Id(Idlocator)));
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }
        public static void ExplicitWaitByCss(string Csslocator, int wait)
        {
            try
            {
                WebDriverWait explicitwait = new WebDriverWait(DriverContext.GetDriver<IWebDriver>(), TimeSpan.FromSeconds(wait));
                explicitwait.PollingInterval = TimeSpan.FromMilliseconds(500);
                explicitwait.IgnoreExceptionTypes(typeof(NoSuchElementException));
                IWebElement element = explicitwait.Until(ExpectedConditions.ElementExists(By.CssSelector(Csslocator)));
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }
        public static void IfElementExistsByXpath(string Xpathlocator, int wait)
        {
            try
            {
                WebDriverWait explicitwait = new WebDriverWait(DriverContext.GetDriver<IWebDriver>(), TimeSpan.FromSeconds(wait));
                explicitwait.PollingInterval = TimeSpan.FromMilliseconds(500);
                explicitwait.IgnoreExceptionTypes(typeof(NoSuchElementException), typeof(ElementNotVisibleException));
                explicitwait.Until(ExpectedConditions.ElementExists(By.XPath(Xpathlocator)));
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }
        //************************end of  WITH WAIT *******************************//
        public static void SelectElementByIndex(By locator, int index)
        {
            select = new SelectElement(GenericHelper.GetElement(locator));
            select.SelectByIndex(index);
        }
        public static void SelectElementByText(By locator, string visibletext)
        {
            select = new SelectElement(GenericHelper.GetElement(locator));
            Thread.Sleep(20000);
            select.SelectByText(visibletext);
        }
        public static void SelectElementByValue(By locator, string valueTexts)
        {
            select = new SelectElement(GenericHelper.GetElement(locator));
            select.SelectByValue(valueTexts);
        }
        public static void SelectElement(IWebElement element, string visibletext)
        {
            select = new SelectElement(element);
            select.SelectByValue(visibletext);
        }
        public static void SelectInDropdownByValue(IWebDriver driver, string selector, string selection)
        {
            var element = new SelectElement(driver.FindElement(By.CssSelector(selector)));
            element.SelectByValue(selection);
        }
        public static void SelectInDropdownByText(IWebDriver driver, string selector, string selection)
        {
            var element = new SelectElement(driver.FindElement(By.CssSelector(selector)));
            element.SelectByText(selection);
        }
        public static void OpenNewTab(IWebDriver driver, string url)
        {
            int oldwindowsize = 0;
            int newwindowsize = 0;
            //get the current window handles /
            string popupHandle = string.Empty;
            ReadOnlyCollection<string> all_windowHandles = driver.WindowHandles;
            //get the current window handles count
            oldwindowsize = all_windowHandles.Count;
            //Open new tabs
            IJavaScriptExecutor js = driver as IJavaScriptExecutor;
            string title = (string)js.ExecuteScript("window.open();");
            //get the new window handles count
            ReadOnlyCollection<string> all_windowHandles_new = driver.WindowHandles;
            newwindowsize = all_windowHandles_new.Count;
            existingWindowHandle = driver.CurrentWindowHandle;
            foreach (string handle in all_windowHandles_new)
            {
                if (handle != existingWindowHandle)
                {
                    popupHandle = handle;
                    break;
                }
            }
            //switch to new window 
            Thread.Sleep(1000);
            driver.SwitchTo().Window(popupHandle);
            Thread.Sleep(15000);
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl(url);
            Thread.Sleep(10000);
        }
        public static Boolean IsElementVisible(IWebDriver driver, string selector)
        {
            try
            {
                FindElementWithXpath(driver, selector);
                return true;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        public static Boolean IsElementNotVisible(IWebDriver driver, string selector)
        {
            try
            {
                FindElementWithXpath(driver, selector);
                return false;
            }
            catch (NoSuchElementException)
            {
                return true;
            }
        }
        public static bool IsElementDisplayed(IWebDriver driver, string selector)
        {
            IWebElement myElement = FindElementWithXpath(driver, selector);
            bool result = myElement.Displayed;
            if (result == true)
            {
                Console.WriteLine("Element is  Displayed");
                return true;
            }
            else
            {
                Console.WriteLine("Element is not   Displayed");
                Assert.Fail("Element is not  Displayed");
                return false;
            }
        }
        public static bool IsElementNotDisplayed(IWebDriver driver, string selector)
        {
            IWebElement myElement = FindElementWithXpath(driver, selector);
            bool result = myElement.Displayed;
            if (result == true)
            {
                Console.WriteLine("Element is  Displayed");
                Assert.Fail("Element is  Displayed");
                return false;
            }
            else
            {
                Console.WriteLine("Element is not Displayed");
                return true;
            }
        }
        public static bool IsElementEnabled(IWebDriver driver, string selector)
        {
            IWebElement myElement = FindElementWithXpath(driver, selector);
            bool result = myElement.Enabled;
            if (result == true)
            {
                Console.WriteLine("Element is  Enabled");
                return true;
            }
            else
            {
                Console.WriteLine("Element is not   Enabled");
                Assert.Fail("Element is not  Enabled");
                return false;
            }
        }
        public static bool IsElementNotEnabled(IWebDriver driver, string selector)
        {
            IWebElement myElement = FindElementWithXpath(driver, selector);
            bool result = myElement.Enabled;
            if (result == true)
            {
                Console.WriteLine("Element is  Enabled");
                Assert.Fail("Element is  Enabled");
                return false;
            }
            else
            {
                Console.WriteLine("Element is not Enabled");
                return true;
            }
        }
        public static bool IsElementEnabled(IWebDriver driver, IWebElement element)
        {
            bool result = element.Enabled;
            if (result == true)
            {
                Console.WriteLine("Element is  Enabled");
                LogHelper.WriteLog(ConfigReader.logFilePath, "Element is  Enabled");
                return true;
            }
            else
            {
                Console.WriteLine("Element is not   Enabled");
                LogHelper.WriteLog(ConfigReader.logFilePath, "Element is not Enabled");
                Assert.Fail("Element is not  Enabled");
                return false;
            }
        }
        public static bool IsElementNotEnabled(IWebDriver driver, IWebElement element)
        {
            bool result = element.Enabled;
            if (result == true)
            {
                Console.WriteLine("Element is  Enabled");
                LogHelper.WriteLog(ConfigReader.logFilePath, "Element is  Enabled");
                Assert.Fail("Element is  Enabled");
                return false;
            }
            else
            {
                Console.WriteLine("Element is not Enabled");
                LogHelper.WriteLog(ConfigReader.logFilePath, "Element is not  Enabled");
                return true;
            }
        }
        //Shwetabh Srivastava----Get Latest Downloaded file from Downloaded folder configured in Config file as per desired capabilities
        public static string getLastDownloadedFile(string folder)
        {
            string latestfile = "";
            var files = new DirectoryInfo(folder).GetFiles("*.*");
            DateTime lastupdated = DateTime.MinValue;
            foreach (FileInfo file in files)
            {
                if (file.LastWriteTime > lastupdated)
                {
                    lastupdated = file.LastWriteTime;
                    latestfile = file.Name;
                }
            }
            Console.Write("LatestFileName: " + latestfile);
            return latestfile;
        }
        public static IReadOnlyCollection<IWebElement> FindElements(IWebDriver driver, string selector)
        {
            return driver.FindElements(By.CssSelector(selector));
        }
        public static IReadOnlyCollection<IWebElement> FindElementsWithXpath(IWebDriver driver, string selector)
        {
            return driver.FindElements(By.XPath(selector));
        }
        public static IWebElement FindElement(IWebDriver driver, string selector)
        {
            return driver.FindElement(By.CssSelector(selector));
        }
        public static IWebElement FindElementWithXpath(IWebDriver driver, string selector)
        {
            return driver.FindElement(By.XPath(selector));
        }
        public static IWebElement FindElementWithXpath( string xpath)
        {
            return DriverContext.GetDriver<IWebDriver>().FindElement(By.XPath(xpath));
        }
        public static IWebElement FindElementWithXpath(By locator)
        {
            return DriverContext.GetDriver<IWebDriver>().FindElement(locator);
        }
        public static string GetValueByElement(IWebDriver driver, string selector)
        {
            return driver.FindElement(By.CssSelector(selector)).Text;
        }
        public static IWebElement Drag(IWebDriver driver, String dragfrom)
        {
            return driver.FindElement(By.CssSelector(dragfrom));
        }
        public static IWebElement Drop(IWebDriver driver, String dropTo)
        {
            return driver.FindElement(By.XPath(String.Format(dropTo)));
        }
        public static bool IsElementVisibleByClass(IWebDriver driver, string className)
        {
            return driver.FindElement(By.ClassName(className)).Displayed;
        }
        public static bool IsElementEnabledByClass(IWebDriver driver, string className)
        {
            return driver.FindElement(By.ClassName(className)).Enabled;
        }
        public static bool IsElementVisibleById(IWebDriver driver, string id)
        {
            return driver.FindElement(By.Id(id)).Displayed;
        }
        public static bool IsElementEnabledById(IWebDriver driver, string id)
        {
            return driver.FindElement(By.Id(id)).Enabled;
        }
        public static bool IsElementEnabledByXpath(IWebDriver driver, string Xpath)
        {
            return driver.FindElement(By.XPath(Xpath)).Enabled;
        }
        public static Boolean SwitchWindows(IWebDriver driver, string title)
        {
            var currentWindow = driver.CurrentWindowHandle;
            List<string> lstWindow = driver.WindowHandles.ToList();
            foreach (string w in lstWindow)
            {
                Console.WriteLine("w:" + w);
                if (w != currentWindow)
                {
                    driver.SwitchTo().Window(w);
                    if (driver.Title == title)
                    {
                        return true;
                    }
                    else
                    {
                        driver.SwitchTo().Window(currentWindow);
                    }
                }
            }
            return false;
        }
        public static bool SwitchToWindowWithTitle(String title)
        {
            IReadOnlyCollection<string> all_windowHandles = DriverContext.GetDriver<IWebDriver>().WindowHandles;
            // Set<String> windowHandles = driver.getWindowHandles();
            foreach (String handle in all_windowHandles)
            {
                DriverContext.GetDriver<IWebDriver>().SwitchTo().Window(handle);
                if (DriverContext.GetDriver<IWebDriver>().Title.Contains(title))
                {
                    break;
                }
            }
            return true;
        }
        //public static void WriteToExcel(string data)
        //{
        //    string myPath = @"C:\Users\ssrivastava4\Documents\Visual Studio 2015\Projects\SHWETABH\SHWETABH\Tellurium-11May-Server\Tellurium\Login.xlsx"; // this must be full path.
        //    FileInfo fi = new FileInfo(myPath);
        //    if (!fi.Exists)
        //    {
        //        Console.Out.WriteLine("file doesn't exists!");
        //    }
        //    else
        //    {
        //        var excelApp = new Microsoft.Office.Interop.Excel.Application();
        //        var workbook = excelApp.Workbooks.Open(myPath);
        //        Worksheet worksheet = workbook.ActiveSheet as Worksheet;
        //        Microsoft.Office.Interop.Excel.Range range = worksheet.Cells[1, 1] as Range;
        //        range.Value2 = data;
        //        //excelApp.Visible = true;
        //        workbook.Save();
        //        workbook.Close();
        //    }
        //}
        public static bool MouseHoverToElement(IWebElement element)
        {
            try
            {
                Thread.Sleep(15000);
                Actions action = new Actions(DriverContext.GetDriver<IWebDriver>());
                action.MoveToElement(element).Build().Perform();
                Thread.Sleep(15000);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return true;
        }
        public static bool MouseHoverToElementbylocator(By locator)
        {
            try
            {
                Thread.Sleep(15000);
                Actions action = new Actions(DriverContext.GetDriver<IWebDriver>());
                action.MoveToElement(DriverContext.GetDriver<IWebDriver>().FindElement(locator)).Build().Perform();
                Thread.Sleep(15000);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return true;
        }
        public static bool MouseHoverToElementandClickbylocator(By locator)
        {
            try
            {
                Thread.Sleep(15000);
                Actions action = new Actions(DriverContext.GetDriver<IWebDriver>());
                action.MoveToElement(DriverContext.GetDriver<IWebDriver>().FindElement(locator)).Click().Build().Perform();
                Thread.Sleep(15000);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return true;
        }
        public static Boolean MouseHoverToElementAndClick(IWebDriver driver, String selector)
        {
            try
            {
                IWebElement we = driver.FindElement(By.XPath(selector));
                Actions action = new Actions(driver);
                action.MoveToElement(we).Build().Perform();
                Thread.Sleep(15000);
                action.MoveToElement(we).Click(we).Build().Perform();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            Console.Write("Mouse Hover to Element and perform click succeed");
            return true;
        }
        public static Boolean MouseHoverToElementAndClick(IWebDriver driver, IWebElement we)
        {
            try
            {
                Actions action = new Actions(driver);
                action.MoveToElement(we).Build().Perform();
                Thread.Sleep(15000);
                action.MoveToElement(we).Click(we).Build().Perform();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            Console.Write("Mouse Hover to Element and perform click succeed");
            return true;
        }
        public static Boolean ScrollToElement(IWebDriver driver, String selector)
        {
            try
            {
                IWebElement we = driver.FindElement(By.CssSelector(selector));
                Actions action = new Actions(driver);
                action.MoveToElement(we).Build().Perform();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            Console.Write("Mouse Hover to Element succeed");
            return true;
        }
        public static void getScreenshot(IWebDriver driver)
        {
            // Start Initializing the variables... 
            Console.WriteLine("Initializing the variables...");
            Console.WriteLine();
            Bitmap memoryImage;
            memoryImage = new Bitmap(1000, 900);
            Size s = new Size(memoryImage.Width, memoryImage.Height);
            // Create the graphics 
            Console.WriteLine("Creating Graphics...");
            Console.WriteLine();
            Graphics memoryGraphics = Graphics.FromImage(memoryImage);
            // Copy data from screen 
            Console.WriteLine("Copying data from screen...");
            Console.WriteLine();
            memoryGraphics.CopyFromScreen(0, 0, 0, 0, s);
            //That's it! Save the image in the directory  
            string str = "";
            try
            {
                //   str = string.Format(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + 
                //        @"\Screenshot.png"); 
                string filename = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString() + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + DateTime.Now.Millisecond.ToString();
                str = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + @"\" + filename + ".png";
            }
            catch (Exception er)
            {
                Console.WriteLine("Sorry, there was an error: " + er.Message);
                Console.WriteLine();
            }
            // Save it! 
            Console.WriteLine("Saving the image...");
            memoryImage.Save(str);
            // Write the message, 
            Console.WriteLine("Picture has been saved...");
            Console.WriteLine();
        }
        // Shwetabh Srivastava---Generates the log file
        public static void WriteLog(IWebDriver driver, String filepath, String strLog)
        {
            StreamWriter log;
            FileStream fileStream = null;
            DirectoryInfo logDirInfo = null;
            FileInfo logFileInfo;
            string logFilePath = filepath;
            logFilePath = logFilePath + partiallogFilePath;
            logFileInfo = new FileInfo(logFilePath);
            logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
            if (!logDirInfo.Exists) logDirInfo.Create();
            if (!logFileInfo.Exists)
            {
                fileStream = logFileInfo.Create();
            }
            else
            {
                fileStream = new FileStream(logFilePath, FileMode.Append);
            }
            log = new StreamWriter(fileStream);
            log.WriteLine(strLog);
            log.Close();
        }
        public static string getExcelSheetName(IWebDriver driver, String filepath, String fileName, String Index)
        {
            String SheetName = null;
            try
            {
                WriteLog(driver, filepath, "In GetSheetName");
                string s = null;
                Workbook wb = new Workbook(filepath + "\\" + fileName);
                int index = Int32.Parse(Index);
                int numberOfSheets = wb.Worksheets.Count;
                WriteLog(driver, filepath, "Number of sheets in this workbook : " + numberOfSheets);
                foreach (Worksheet worksheet in wb.Worksheets)
                {
                    if (worksheet.Index == (index))
                    {
                        SheetName = worksheet.Name;
                        WriteLog(driver, filepath, "Sheet name for Index " + Index + "  is  " + SheetName);
                    }
                }
            }
            catch (Exception e)
            {
                WriteLog(driver, filepath, " getExcelSheetName exception :" + e.Message);
            }
            return SheetName;
        }
    }
}
