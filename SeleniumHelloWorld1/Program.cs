using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Text;
using System.Threading;

namespace SeleniumHelloWorld
{
    internal class Program
    {
        private const int QUIT_DELEY_SEC = 10;
        private const string URL_BING = @"https://www.bing.com/";
        private const string URL_GOV_US = @"https://usr.minjust.gov.ua/ua/freesearch/";
        private const string ID_SEARCH_INPUT = "query";
        private const string NO = "3403201375";
        private const string FRAME_CLASS_NAME = "extiframe";
        private const string SCREEN_MAX = "--start-maximized";

        private static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static void Run()
        {
            ChromeOptions chromeOpt = new ChromeOptions();
            chromeOpt.AddArgument(SCREEN_MAX);
            IWebDriver driver = new ChromeDriver(chromeOpt);

            driver.Navigate().GoToUrl(URL_GOV_US);
            driver.SwitchTo().Frame(driver.FindElement(By.ClassName(FRAME_CLASS_NAME)));
            var searchCtrl = WaitForElementToAppear(driver, 5, By.Id(ID_SEARCH_INPUT));
            searchCtrl.SendKeys(NO);
            searchCtrl.SendKeys(Keys.Enter);


            Thread.Sleep(TimeSpan.FromSeconds(10));
            WaitForElementToAppear(driver, 30, By.ClassName("searchother"));

            //driver.SwitchTo().Frame(driver.FindElement(By.ClassName(FRAME_CLASS_NAME)));
            var table = WaitForElementToAppear(driver, 30, By.Id("detailtable"));
            //var frame = WaitForElementToAppear(driver, 30, By.ClassName(FRAME_CLASS_NAME));
            var body = WaitForElementToAppear(driver, 30, By.TagName("body"));

            string  data = table.GetAttribute("innerHTML");
            data = body.GetAttribute("innerHTML");
            string codeBase = AppDomain.CurrentDomain.BaseDirectory;
            string name = $"{NO}.html";
            string output = Path.Combine(codeBase, name);
            using (TextWriter writer = File.CreateText(output))
            {
                writer.WriteLine(data);
            }
            driver.Quit();
        }

        public static IWebElement WaitForElementToAppear(IWebDriver driver, int waitTime, By waitingElement)
        {
            IWebElement wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitTime)).Until(ExpectedConditions.ElementExists(waitingElement));
            return wait;
        }
    }
}