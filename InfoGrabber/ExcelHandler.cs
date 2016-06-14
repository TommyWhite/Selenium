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
    }
}