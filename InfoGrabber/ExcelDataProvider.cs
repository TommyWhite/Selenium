using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;

namespace EmployeeInfoGrabber
{
    public class ExcelDataProvider
    {
        private string GetConnectionString(string fullFilePath)
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
            string connectionString = GetConnectionString(fullFilePath);
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
    }
}