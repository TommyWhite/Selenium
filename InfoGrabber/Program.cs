using System;
using System.IO;

namespace EmployeeInfoGrabber
{
    internal class Program
    {
        private static void Main(string[] args)
        {

            DataGrabber grabber = new DataGrabber();

            string fileName = "3.xlsx";
            var excelFile = $@"C:\Users\artemm\Desktop\EmployeeInfoGrabber\InfoGrabber\bin\Debug\{fileName}";

            var outputPath = Path.Combine(AppContext.BaseDirectory, "reports");
            grabber.Run(excelFile, outputPath);
        }
    }
}