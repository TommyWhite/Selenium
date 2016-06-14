using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
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
            GlobalVars global = new GlobalVars();

            string dataFile = ConfigurationManager.AppSettings["excelWithTaxNumbers"];
            if (string.IsNullOrEmpty(dataFile) && File.Exists(dataFile))
            {
                throw new FileNotFoundException("Excel file with input data is not found!");
            }

            var ddt = new ExcelDataProvider();
            string fileName = "3.xlsx";
            var filePath = $@"C:\Users\artemm\Desktop\EmployeeInfoGrabber\InfoGrabber\bin\Debug\{fileName}";
            var data = ddt.ReadExcelFile(filePath);

            List<string> list = new List<string>();

            foreach (DataRow row in data.Tables[0].Rows)
            {
                var number = row.ItemArray.Select(NO => NO.ToString()).ToList();
                list.AddRange(number);
            }

            DataGrabber grabber = new DataGrabber();

            try
            {
                foreach (var item in list)
                {
                    string codeBase = global.BaseDir;
                    string name = $"{item}.html";
                    var grabbed = grabber.GrabData(item);
                    grabber.SaveDataTo(codeBase, name, grabbed);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                grabber.Dispose();
            }
        }
    }

    
}