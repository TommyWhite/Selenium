using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EmployeeInfoGrabber;

namespace UnitTests
{
    [TestClass]
    public class ExcelHandlerTests
    {
        private ExcelHandler xlsHandler = new ExcelHandler();

        [TestMethod]
        public void ReadExcelFile()
        {
            var ds = xlsHandler.ReadExcelFile("MOCK_DATA.xlsx");
            bool dsIsNotEmpty = ds.Tables[0].Rows.Count != 0;
            Assert.IsTrue(dsIsNotEmpty);
        }

        [TestMethod]
        public void WriteExcelFile()
        {

        }

    }
}
