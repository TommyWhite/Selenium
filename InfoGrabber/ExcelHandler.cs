using ExcelLibrary;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;

namespace EmployeeInfoGrabber
{
    public class ExcelHandler
    {
        private Microsoft.Office.Interop.Excel.Application oXL;
        private Microsoft.Office.Interop.Excel._Workbook oWB;
        private Microsoft.Office.Interop.Excel._Worksheet oSheet;
        private Microsoft.Office.Interop.Excel.Range oRng;

        private string BuildConnectionString(string fullFilePath)
        {
            Dictionary<string, string> props = new Dictionary<string, string>()
            {
                ["Provider"] = "Microsoft.ACE.OLEDB.12.0;",
                ["Extended Properties"] = "Excel 12.0 XML",
                ["Data Source"] = fullFilePath
            };
            StringBuilder sb = new StringBuilder();
            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append($"{prop.Key}={prop.Value};");
            }

            return sb.ToString();
        }

        public DataSet ReadExcelFile(string fullFilePath, string sheetName = "Sheet1$")
        {
            DataSet ds = new DataSet();
            string connectionString = BuildConnectionString(fullFilePath);
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand()
                {
                    Connection = conn,
                    CommandText = $"SELECT * FROM [{sheetName}]"
                };

                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
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

        public void WriteExcelFile(string fullNamePath, DataSet data)
        { 
            DataSetHelper.CreateWorkbook(fullNamePath, data);
        }

        public void CreateXlsDoc(string outputFile, string rangeFrom, string rangeTo, string[,] data)
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

            oSheet.get_Range("A1", "P1").Font.Bold = true;
            oSheet.Columns["A:P"].VerticalAlignment =
                Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            oSheet.Columns["A:P"].WrapText = true;
            oSheet.get_Range("A1", "P1").ColumnWidth = 20;
            oSheet.get_Range(rangeFrom, rangeTo).Value2 = data;
            oRng = oSheet.get_Range("A1", "D1");
            oRng.EntireColumn.AutoFit();

            oXL.UserControl = false;

            oWB.SaveAs(outputFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            oWB.Close();
        }
    }
}