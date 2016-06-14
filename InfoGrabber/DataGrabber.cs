using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading;

namespace EmployeeInfoGrabber
{
    public class DataGrabber : IDisposable
    {
        private bool _isDisposed = false;

        private static IWebDriver driver;

        private void Initialize()
        {
            ChromeOptions chromeOpt = new ChromeOptions();
            chromeOpt.AddArgument(GlobalVars.SCREEN_MAX);
            driver = new ChromeDriver(chromeOpt);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_isDisposed)
            {
                driver.Quit();
                driver.Dispose();
                if (disposing)
                {
                    driver = null;
                }
            }
        }

        public DataGrabber()
        {
            Initialize();
        }

        public string GrabData(string taxNum)
        {
            driver.Navigate().GoToUrl(GlobalVars.URL_GOV_US);
            driver.SwitchTo().Frame(driver.FindElement(By.ClassName(GlobalVars.FRAME_CLASS_NAME)));
            var searchField = WaitForElementToAppear(driver, 5, By.Id(GlobalVars.ID_SEARCH_INPUT));
            searchField.SendKeys(GlobalVars.TAX_NUMBER);
            searchField.SendKeys(Keys.Enter);

            Thread.Sleep(TimeSpan.FromSeconds(30));
            //TODO: W8 and click to the OK button on the reCAPTCHA
            WaitForElementToAppear(driver, 90, By.ClassName("searchother"));

            string edittedContent;
            try
            {
                var iframeCtrl = WaitForElementToAppear(driver, 90, By.TagName("html"));
                string frameContent = iframeCtrl.GetAttribute("innerHTML");
                string removableString = "<form method=\"post\" class=\"searchother\"><input name=\"searchother\" type=\"submit\" value=\"Шукати ще\"></form>";
                edittedContent = frameContent.Replace(removableString, "");
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to grab content from html page. Message: {ex.Message}");
            }

            return edittedContent;
        }

        public IWebElement WaitForElementToAppear(IWebDriver driver, int waitTime, By waitingElement)
        {
            IWebElement wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitTime)).Until(ExpectedConditions.ElementExists(waitingElement));
            return wait;
        }

        public void SaveDataTo(string baseFolder, string fileNameWithExt, string data)
        {
            string output = Path.Combine(baseFolder, fileNameWithExt);
            using (TextWriter writer = File.CreateText(output))
            {
                writer.WriteLine(data);
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        //TODO: Implement input file and output dir for HTML reports.
        public void Run(string inputXML, string outputDir)
        {
            var ddt = new ExcelDataProvider();
            var data = ddt.ReadExcelFile(inputXML);
            List<string> VATINList = new List<string>();

            foreach (DataRow row in data.Tables[0].Rows)
            {
                var number = row.ItemArray.Select(NO => NO.ToString()).ToList();
                VATINList.AddRange(number);
            }

            try
            {
                foreach (var taxNumber in VATINList)
                {
                    string itemName = $"{taxNumber}.html";
                    var grabbed = GrabData(taxNumber);
                    SaveDataTo(outputDir, itemName, grabbed);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Dispose();
            }
        }
    }
}