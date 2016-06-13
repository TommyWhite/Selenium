using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
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

            //driver.SwitchTo().Frame(driver.FindElement(By.ClassName(FRAME_CLASS_NAME)));
            var table = WaitForElementToAppear(driver, 30, By.Id("detailtable"));
            //var frame = WaitForElementToAppear(driver, 30, By.ClassName(FRAME_CLASS_NAME));
            var body = WaitForElementToAppear(driver, 30, By.TagName("body"));

            //TODO: Chose the right data to be saved.
            string tableInner = table.GetAttribute("innerHTML");
            string bodyInner = body.GetAttribute("innerHTML");

            return tableInner;
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
    }

    public class ExcelDataProvider
    {
        private string GetConnectionString(string fileName)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            props["Extended Properties"] = "Excel 12.0 XML";
            props["Data Source"] = fileName;

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        public DataSet ReadExcelFile(string filePath)
        {
            DataSet ds = new DataSet();

            string connectionString = GetConnectionString(filePath);

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                string sheetName = "Sheet1$";

                cmd.CommandText = $"SELECT * FROM [{sheetName}]";

                DataTable dt = new DataTable()
                {
                    TableName = sheetName
                };

                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);

                ds.Tables.Add(dt);

                conn.Close();
            }

            return ds;
        }
    }
}