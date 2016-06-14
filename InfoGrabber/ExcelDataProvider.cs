using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Text;

namespace EmployeeInfoGrabber
{
    public class ExcelDataProvider
    {
        private string GetConnectionString(string fileName)
        {
            Dictionary<string, string> props = new Dictionary<string, string>();

            props["Provider"] = "Microsoft.ACE.OLEDB.12.0;";
            props["Extended Properties"] = "Excel 12.0 XML";
            props["Data Source"] = fileName;

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }

        public DataSet ReadExcelFile(string filePath)
        {
            DataSet ds = new DataSet();
            string connectionString = GetConnectionString(filePath);
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
