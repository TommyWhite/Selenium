using ExcelLibrary;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;

namespace EmployeeInfoGrabber
{
    public class ExcelHandler
    {
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

        public DataSet ReadExcelFile(string fullFilePath)
        {
            DataSet ds = new DataSet();
            string connectionString = BuildConnectionString(fullFilePath);
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;

                // Get all Sheets in Excel File
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string sheetName = "Sheet1$";
                cmd.CommandText = $"SELECT * FROM [{sheetName}]";
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
            //TODO: Test and remove.
            fullNamePath = @"MyExcelFile.xls";
            #region Just for testion purposes 
            //Create the data set and table
            DataSet ds = new DataSet("New_DataSet")
            {
                Locale = System.Threading.Thread.CurrentThread.CurrentCulture
            };
            DataTable dt = new DataTable("New_DataTable")
            {
                Locale = System.Threading.Thread.CurrentThread.CurrentCulture
            };
            #endregion
            
            DataSetHelper.CreateWorkbook(fullNamePath, ds);
        }
    }
}