using EmployeeInfoGrabber;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace UnitTests.PrototypeCodeSnippets
{
    [TestClass]
    public class XlsManipulationTest
    {
        private Microsoft.Office.Interop.Excel.Application oXL;
        private Microsoft.Office.Interop.Excel._Workbook oWB;
        private Microsoft.Office.Interop.Excel._Worksheet oSheet;
        private Microsoft.Office.Interop.Excel.Range oRng;

        [TestMethod]
        public void PrototypeCreateXlsDoc()
        {
            //Start Excel and get Application object.
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;

            //Get a new workbook.
            oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

            //Add table headers going cell by cell.
            oSheet.Cells[1, 1] = @"Прізвище, ім'я, по батькові фізичної особи";
            oSheet.Cells[1, 2] = @"Місце проживання";
            oSheet.Cells[1, 3] = @"Види діяльності";
            oSheet.Cells[1, 4] = @"Дата державної реєстрації, дата та номер запису в Єдиному державному реєстрі про включення до Єдиного державного реєстру відомостей про фізичну особу-підприємця – у разі, коли державна реєстрація фізичної особи-підприємця була проведена до набрання чинності Законом України “Про державну реєстрацію юридичних осіб та фізичних осіб-підприємців”";
            oSheet.Cells[1, 5] = @"Дата та номер запису про проведення державної реєстрації фізичної особи-підприємця";
            oSheet.Cells[1, 6] = @"Місцезнаходження реєстраційної справи";
            oSheet.Cells[1, 7] = @"Дата та номер запису про взяття та зняття з обліку, назва та ідентифікаційні коди органів статистики, Міндоходів, Пенсійного фонду України, в яких фізична особа-підприємець перебуває на обліку:";
            oSheet.Cells[1, 8] = @"Дані органів державної статистики про основний вид економічної діяльності фізичної особи-підприємця, визначений на підставі даних державних статистичних спостережень відповідно до статистичної методології за підсумками діяльності за рік";
            oSheet.Cells[1, 9] = @"Дані про реєстраційний номер платника єдиного внеску, клас професійного ризику виробництва платника єдиного внеску за основним видом його економічної діяльності";
            oSheet.Cells[1, 10] = @"Термін, до якого фізична особа-підприємець перебуває на обліку в органі Міндоходів за місцем попередньої реєстрації, у разі зміни місця проживання фізичної особи-підприємця";
            oSheet.Cells[1, 11] = @"Дані про перебування фізичної особи-підприємця в процесі припинення підприємницької діяльності, банкрутства";
            oSheet.Cells[1, 12] = @"Прізвище, ім'я, по батькові особи, яка призначена управителем майна фізичної особи-підприємця";
            oSheet.Cells[1, 13] = @"Дата та номер запису про державну реєстрацію припинення підприємницької діяльності фізичною особою-підприємцем, підстава для його внесення";
            oSheet.Cells[1, 14] = @"Дата відміни державної реєстрації припинення підприємницької діяльності фізичною особою-підприємцем, підстава її внесення";
            oSheet.Cells[1, 15] = @"Дата відкриття виконавчого провадження щодо фізичної особи - підприємця (для незавершених виконавчих проваджень)";
            oSheet.Cells[1, 16] = @"Інформація про здійснення зв'язку з фізичною особою-підприємцем";

            //Format A1:D1 as bold, vertical alignment = center.
            oSheet.get_Range("A1", "P1").Font.Bold = true;
            oSheet.Columns["A:P"].VerticalAlignment =
                Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            oSheet.Columns["A:P"].WrapText = true;
            oSheet.get_Range("A1", "P1").ColumnWidth = 20;

            // Create an array to multiple values at once.
            //string[,] saNames = new string[5, 2];

            //saNames[0, 0] = "John";
            //saNames[0, 1] = "Smith";
            //saNames[1, 0] = "Tom";
            //saNames[4, 1] = "Johnson";

            ////Fill A2:B6 with an array of values (First and Last Names).
            //oSheet.get_Range("A2", "B6").Value2 = saNames;

            ////Fill C2:C6 with a relative formula (=A2 & " " & B2).
            //oRng = oSheet.get_Range("C2", "C6");
            //oRng.Formula = ("=A2 & \" \" & B2");

            ////Fill D2:D6 with a formula(=RAND()*100000) and apply format.
            //oRng = oSheet.get_Range("D2", "D6");
            //oRng.Formula = "=RAND()*100000";
            //oRng.NumberFormat = "$0.00";

            //AutoFit columns A:D.tt
            oRng = oSheet.get_Range("A1", "D1");
            oRng.EntireColumn.AutoFit();

            oXL.UserControl = false;

            string codeBase = AppContext.BaseDirectory;
            int fileName = 0;
            string outputFile;
            do
            {
                outputFile = Path.Combine(codeBase, $"{fileName++}.xlsx");
            } while (File.Exists(outputFile));

            oWB.SaveAs(outputFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            oWB.Close();
        }

        [TestMethod]
        public void CreateXlsDoc()
        {
            ExcelHandler excelHndlr = new ExcelHandler();

            string codeBase = AppContext.BaseDirectory;
            int fileName = 0;
            string outputFile;
            do
            {
                outputFile = Path.Combine(codeBase, $"{fileName++}.xlsx");
            } while (File.Exists(outputFile));

            const int ASCII_ALPHAS_OFFSET = 65;
            string[,] data = new string[10, 16];
            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    data[i, j] = Convert.ToChar(ASCII_ALPHAS_OFFSET + j).ToString();
                }
            }

            excelHndlr.CreateXlsDoc(outputFile, $"A{2}", $"P{10}", data);
        }

        private ExcelHandler xlsHandler = new ExcelHandler();

        [TestMethod]
        public void ReadExcelFile()
        {
            var ds = xlsHandler.ReadExcelFile("MOCK_DATA.xlsx", "data$");
            bool dsIsNotEmpty = ds.Tables[0].Rows.Count != 0;
            Assert.IsTrue(dsIsNotEmpty, "Fails to read xml file.");
        }
    }
}