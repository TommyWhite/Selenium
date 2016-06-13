using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Threading;

namespace SeleniumHelloWorld
{
    public class DataGrabber : IDisposable
    {
        private bool _isDisposed = false;

        private static IWebDriver driver;

        private void Initialize()
        {
            ChromeOptions chromeOpt = new ChromeOptions();
            chromeOpt.AddArgument(Globals.SCREEN_MAX);
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

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public static void GrabData(string taxNum)
        {
            driver.Navigate().GoToUrl(Globals.URL_GOV_US);
            driver.SwitchTo().Frame(driver.FindElement(By.ClassName(Globals.FRAME_CLASS_NAME)));
            var searchField = WaitForElementToAppear(driver, 5, By.Id(Globals.ID_SEARCH_INPUT));
            searchField.SendKeys(Globals.TAX_NUMBER);
            searchField.SendKeys(Keys.Enter);

            Thread.Sleep(TimeSpan.FromSeconds(30));
            //TODO: W8 and click to the OK button on the reCAPTCHA
            WaitForElementToAppear(driver, 90, By.ClassName("searchother"));

            //driver.SwitchTo().Frame(driver.FindElement(By.ClassName(FRAME_CLASS_NAME)));
            var table = WaitForElementToAppear(driver, 30, By.Id("detailtable"));
            //var frame = WaitForElementToAppear(driver, 30, By.ClassName(FRAME_CLASS_NAME));
            var body = WaitForElementToAppear(driver, 30, By.TagName("body"));

            string data = table.GetAttribute("innerHTML");
            data = body.GetAttribute("innerHTML");
            string codeBase = AppDomain.CurrentDomain.BaseDirectory;
            string name = $"{Globals.TAX_NUMBER}.html";
            SaveDataTo(codeBase, name, data);

            driver.Quit();
        }

        public static IWebElement WaitForElementToAppear(IWebDriver driver, int waitTime, By waitingElement)
        {
            IWebElement wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitTime)).Until(ExpectedConditions.ElementExists(waitingElement));
            return wait;
        }

        public static void SaveDataTo(string baseFolder, string fileNameWithExt, string data)
        {
            string output = Path.Combine(baseFolder, fileNameWithExt);
            using (TextWriter writer = File.CreateText(output))
            {
                writer.WriteLine(data);
            }
        }

    }
}
