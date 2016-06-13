using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Configuration;
using System.IO;
using System.Threading;

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
            var settings = ConfigurationManager.GetSection(typeof(InputData).Name);
            InputData inputData = settings as InputData;

            if (inputData != null)
            { 
                string file = inputData.ExcelFile;
            }
            else
            {
                throw new FileNotFoundException("Excel file with input data is not found!");
            }

            DataGrabber grabber = new DataGrabber();
            try
            {

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