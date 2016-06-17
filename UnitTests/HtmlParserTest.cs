using EmployeeInfoGrabber;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;

namespace UnitTests
{
    [TestClass]
    public class HtmlParserTest
    {
        [TestMethod]
        public void ConternParsing()
        {
            HtmlHandler parser = new HtmlHandler();

            string content;
            using (StreamReader reader = new StreamReader("index.html"))
            {
                content = reader.ReadToEnd();
            }

            var data = parser.ParseHtml(content);
            Assert.IsTrue(data.Length == 32, "Empty parsed data.");
        }

        [TestMethod]
        public void CollectingAllReportsTest()
        {
            HtmlHandler parser = new HtmlHandler();
            var data = parser.ReadReportData(".");
        }
    }
}