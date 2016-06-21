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
            searchField.SendKeys(taxNum);
            searchField.SendKeys(Keys.Enter);
            Thread.Sleep(TimeSpan.FromSeconds(1));

            driver.SwitchTo().Frame(driver.FindElement(By.TagName("iframe")));
            const string CAPTCHA = "recaptcha-anchor-label";
            var captchaCtrl = driver.FindElement(By.Id(CAPTCHA));
            captchaCtrl.Click();

            bool waitableCtrlExists = false;
            do
            {
                Thread.Sleep(TimeSpan.FromSeconds(5));
                IWebElement waitableCtrl = null;
                try
                {
                    driver.SwitchTo().DefaultContent();
                    driver.SwitchTo().Frame(driver.FindElement(By.TagName("iframe")));
                    var btnOK = WaitForElementToAppear(driver, 15, By.CssSelector("[value='OK']"));
                    btnOK.Click();

                    driver.SwitchTo().DefaultContent();
                    driver.SwitchTo().Frame(driver.FindElements(By.TagName("iframe"))[0]);
                    waitableCtrl = WaitForElementToAppear(driver, 1, By.ClassName("detailinfo"));

                }
                catch (Exception ex)
                {
                    /*
                     * Exception caught if captcha is not passed in 5 seconds.
                     * Just pending flow while needed control is not found.
                     * Supposed that user is passing reCAPTCHA
                    */
                }
                finally
                {
                    waitableCtrlExists = waitableCtrl == null ? true : false;
                }
            } while (waitableCtrlExists);

            
            
            var queryResults = driver.FindElements(By.TagName("tr"));
            var buttonDetails = queryResults.Where(element => element.GetAttribute("innerHTML")
            .Contains("не перебуває в процесі припинення"))
            .Select(f => f.FindElements(By.TagName("input")).Last()).FirstOrDefault();
            buttonDetails.Click();
            Thread.Sleep(TimeSpan.FromSeconds(1));

            IWebElement contentWaitableCtrl = WaitForElementToAppear(driver, 10, By.ClassName("searchother"));
            var iframeCtrl = WaitForElementToAppear(driver, 90, By.TagName("html"));
            string frameContent = iframeCtrl.GetAttribute("innerHTML");
            string removableString = "<form method=\"post\" class=\"searchother\"><input name=\"searchother\" type=\"submit\" value=\"Шукати ще\"></form>";
            string edittedContent = frameContent.Replace(removableString, "");

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

        public void Run(string inputXML, string outputDir)
        {
            ExcelHandler ddt = new ExcelHandler();
            DataSet ds = ddt.ReadExcelFile(inputXML);
            List<string> VATINList = new List<string>();
            var table = ds.Tables[0];
            for (int i = 0; i < table.Rows.Count; i++)
            {
                var data = table.Rows[i][0].ToString();
                var outputFile = Path.Combine(outputDir, $"{data}.html");
                if (string.IsNullOrEmpty(data) || File.Exists(outputFile) || VATINList.Contains(data))
                    continue;
                else
                    VATINList.Add(data);
            }

            foreach (var taxNumber in VATINList)
            {
                string fileName = $"{taxNumber}.html";
                var grabbed = GrabData(taxNumber);
                SaveDataTo(outputDir, fileName, grabbed);
            }

            ddt.Dispose();
        }
    }
}