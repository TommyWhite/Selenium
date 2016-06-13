using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace SeleniumHelloWorld
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Initialize data via Configuration Section (need to be fixed)
            //var settings = ConfigurationManager.GetSection(typeof(InputData).Name);
            //InputData inputData = settings as InputData;
            //string dataFile;
            //if (inputData != null)
            //{
            //    dataFile = inputData.ExcelFile;
            //}
            //else
            //{
            //    throw new FileNotFoundException("Excel file with input data is not found!");
            //}

            string dataFile = ConfigurationManager.AppSettings["excelWithTaxNumbers"];
            if (string.IsNullOrEmpty(dataFile) && File.Exists(dataFile))
            {
                throw new FileNotFoundException("Excel file with input data is not found!");
            }

            var ddt = new ExcelDataProvider();
            string fileName = "3.xlsx";
            var filePath = $@"C:\Users\artemm\Desktop\Selenium\SeleniumHelloWorld1\bin\Debug\{fileName}";
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
                    string codeBase = AppDomain.CurrentDomain.BaseDirectory;
                    string name = $"{Globals.TAX_NUMBER}.html";
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