using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;

namespace EmployeeInfoGrabber
{
    public class ExcelHandler : IDisposable
    {
        private bool _isDisposed;

        private Microsoft.Office.Interop.Excel.Application oXL;
        private Microsoft.Office.Interop.Excel._Workbook oWB;
        private Microsoft.Office.Interop.Excel._Worksheet oSheet;
        private Microsoft.Office.Interop.Excel.Range oRng;

        /// <summary>
        /// Gets connection string (x86)
        /// </summary>
        /// <param name="fullFilePath">File to read with OleDbConnection</param>
        /// <returns>Connection string with specified data source.</returns>
        private string BuildConnectionString(string fullFilePath)
        {
            return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={fullFilePath};Extended Properties=\"Excel 12.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text\"";
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

        public void CreateXlsDoc(string outputFile, string rangeFrom, string rangeTo, string[,] data)
        {
            //Start Excel and get Application object.
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = false;

            //Get a new workbook.
            oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

            //Add table headers going cell by cell.
            List<string> columns = new List<string>()
            {
                @"Ідентифікаційний номер",
                @"Прізвище, ім'я, по батькові фізичної особи",
                @"Місце проживання",
                @"Види діяльності",
                @"Дата державної реєстрації, дата та номер запису в Єдиному державному реєстрі про включення до Єдиного державного реєстру відомостей про фізичну особу-підприємця – у разі, коли державна реєстрація фізичної особи-підприємця була проведена до набрання чинності Законом України “Про державну реєстрацію юридичних осіб та фізичних осіб-підприємців”",
                @"Дата та номер запису про проведення державної реєстрації фізичної особи-підприємця",
                @"Місцезнаходження реєстраційної справи",
                @"Дата та номер запису про взяття та зняття з обліку, назва та ідентифікаційні коди органів статистики, Міндоходів, Пенсійного фонду України, в яких фізична особа-підприємець перебуває на обліку:",
                @"Дані органів державної статистики про основний вид економічної діяльності фізичної особи-підприємця, визначений на підставі даних державних статистичних спостережень відповідно до статистичної методології за підсумками діяльності за рік",
                @"Дані про реєстраційний номер платника єдиного внеску, клас професійного ризику виробництва платника єдиного внеску за основним видом його економічної діяльності",
                @"Термін, до якого фізична особа-підприємець перебуває на обліку в органі Міндоходів за місцем попередньої реєстрації, у разі зміни місця проживання фізичної особи-підприємця",
                @"Дані про перебування фізичної особи-підприємця в процесі припинення підприємницької діяльності, банкрутства",
                @"Прізвище, ім'я, по батькові особи, яка призначена управителем майна фізичної особи-підприємця",
                @"Дата та номер запису про державну реєстрацію припинення підприємницької діяльності фізичною особою-підприємцем, підстава для його внесення",
                @"Дата відміни державної реєстрації припинення підприємницької діяльності фізичною особою-підприємцем, підстава її внесення",
                @"Дата відкриття виконавчого провадження щодо фізичної особи - підприємця (для незавершених виконавчих проваджень)",
                @"Інформація про здійснення зв'язку з фізичною особою-підприємцем",
            };

            const int HEADER_ROW_INDEX = 1;
            for (int i = 0; i < columns.Count; i++)
            {
                oSheet.Cells[HEADER_ROW_INDEX, i + 1] = columns[i];
            }

            const int TITLE_OFFSET = 2;
            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    oSheet.Cells[i + TITLE_OFFSET, j + 1] = data[i, j];
                }
            }
            oSheet.get_Range("A1", "Q1").Font.Bold = true;
            oSheet.Columns["A:Q"].VerticalAlignment =
                Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            oSheet.Columns["A:Q"].WrapText = true;
            oSheet.get_Range("A1", "Q1").ColumnWidth = 20;
            oSheet.get_Range(rangeFrom, rangeTo).Value2 = data;
            for (int i = 0; i < data.GetLength(0); i++)
            {
                string cellWithRef = "A" + (i + TITLE_OFFSET);
                Microsoft.Office.Interop.Excel.Range excelCell = oSheet.get_Range(cellWithRef, Type.Missing);
                Match match = Regex.Match(excelCell.Value2.ToString(), @"\d+(.html)$");
                string linkTitle = match.Groups[0].Value.Replace(".html", "");
                excelCell.Hyperlinks.Add(excelCell, excelCell.Value2.ToString(), Type.Missing, "Delphi LLC", linkTitle);
            }
            oRng = oSheet.get_Range("A1", "Q1");
            oRng.EntireColumn.AutoFit();

            oXL.UserControl = false;
            string dt = DateTime.Now.ToShortDateString();
            outputFile = Path.Combine(outputFile, $"Report_{dt}.xlsx");
            oWB.SaveAs(outputFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            oWB.Close();
        }

        public void Dispose()
        {
            Dispose(false);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_isDisposed)
            {
                if (disposing)
                {
                    oXL = null;
                    oWB = null;
                    oSheet = null;
                    oRng = null;
                }
                var ps = Process.GetProcessesByName("EXCEL");
                foreach (Process proc in ps)
                {
                    proc.Kill();
                }

                _isDisposed = true;
            }
        }
    }
}
