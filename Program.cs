using ExcelDataReader;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace ReadExcelWriteSqlServerConsoleApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string excelFilePath = "tutorial.xlsx";         //"/path/to/your/excel/file.xls";
            string connectionString = GetConnectionString();//"your-sql-server-connection-string";

            DataTable table = ReadExcel(excelFilePath);

            // 刪除第一行（表頭）
            if (table.Rows.Count > 0)
            {
                table.Rows.RemoveAt(0);
            }

            foreach (DataRow row in table.Rows)
            {
                // 假設日期數據在第一列
                if (DateTime.TryParse(row[3].ToString(), out DateTime dateValue))
                {
                    Console.WriteLine($"日期: {dateValue.ToShortDateString()}");
                    row[3] = dateValue.ToString("yyyy-MM-dd HH:mm:ss");
                }
                else
                {
                    Console.WriteLine("無效的日期格式");
                }
            }

            WriteToSqlServer(table, connectionString);
        }

        public static DataTable ReadExcel(string filePath)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var result = reader.AsDataSet();
            return result.Tables[0];
        }

        public static void WriteToSqlServer(DataTable table, string connectionString)
        {
            using var bulkCopy = new SqlBulkCopy(connectionString);
            bulkCopy.DestinationTableName = "tutorial"; //"YourTableName";
            bulkCopy.WriteToServer(table);
        }

        private static string GetConnectionString()
        // To avoid storing the sourceConnection string in your code,
        // you can retrieve it from a configuration file.
        {
            //return "Data Source=(local); " +
            //    " Integrated Security=true;" +
            //    "Initial Catalog=AdventureWorks;";

            return "Persist Security Info=False;User ID=*****;Password=*****;Initial Catalog=AdventureWorks;Server=****-**-**\\MSSQL2022";
        }
    }
}
