using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System;
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
    
        private static void Main(string[] args)
        {
            try
            {
                //asdfasdfasdfasdf
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
            chromeOpt.AddArgument("--start-maximized");
            IWebDriver driver = new ChromeDriver(chromeOpt);

            driver.Navigate().GoToUrl(URL_GOV_US);
            driver.SwitchTo().Frame(driver.FindElement(By.ClassName("extiframe")));
            var searchCtrl = WaitForElementToAppear(driver, 5, By.Id(ID_SEARCH_INPUT));
            searchCtrl.SendKeys(NO);
            searchCtrl.SendKeys(Keys.Enter);
            
            Thread.Sleep(15000);
            //driver.SwitchTo().Frame(driver.FindElement(By.ClassName("extiframe")));
            var table = WaitForElementToAppear(driver, 30, By.Id("detailtable"));
            var ctrls = driver.FindElements(By.TagName("td"));
            foreach (var item in ctrls)
            {
                Console.WriteLine(item.Text);
            }
            //driver.Quit();
        }

        public static IWebElement WaitForElementToAppear(IWebDriver driver, int waitTime, By waitingElement)
        {
            IWebElement wait = new WebDriverWait(driver, TimeSpan.FromSeconds(waitTime)).Until(ExpectedConditions.ElementExists(waitingElement));
            return wait;
        }
    }
}