using EmployeeInfoGrabber;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests
{
    [TestClass]
    public class ExcelHandlerTests
    {
        private ExcelHandler xlsHandler = new ExcelHandler();

        [TestMethod]
        public void ReadExcelFile()
        {
            var ds = xlsHandler.ReadExcelFile("MOCK_DATA.xlsx", "data$");
            bool dsIsNotEmpty = ds.Tables[0].Rows.Count != 0;
            Assert.IsTrue(dsIsNotEmpty, "Fails to read xml file.");
        }

        [TestMethod]
        public void WriteExcelFile()
        {
            var ds = xlsHandler.ReadExcelFile("MOCK_DATA.xls", "data$");
            xlsHandler.WriteExcelFile("TEST_GENERATED_FILE.xml", ds);
            
            bool dsIsNotEmpty = ds.Tables[0].Rows.Count != 0;
            Assert.IsTrue(dsIsNotEmpty, "Fails to read xml file.");
        }
    }
}