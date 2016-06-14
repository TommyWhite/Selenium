using System;
using System.IO;

namespace EmployeeInfoGrabber
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            string fileName = "input.xlsx";
            string inputFullPath = Path.Combine(AppContext.BaseDirectory, "Resources", "Input", fileName);
            string outputFullPath = Path.Combine(AppContext.BaseDirectory, "Resources", "Output");

            DataGrabber grabber = new DataGrabber();
            grabber.Run(inputFullPath, outputFullPath);
        }
    }
}