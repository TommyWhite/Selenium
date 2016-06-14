using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace EmployeeInfoGrabber
{
    internal class Program
    {
        private static void Main(string[] args)
        {

            DataGrabber grabber = new DataGrabber();
            grabber.Run(null, null);
        }
    }
}