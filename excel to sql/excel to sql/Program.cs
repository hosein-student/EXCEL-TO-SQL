using ClosedXML.Excel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace ExcelToSql
{
    class Program
    {






        static void Main()
        {
            string excelPath = @"E:\Book1.xlsx"; // ← مسیر فایل اکسل رو دقیق وارد کن
            string connectionString = "Data Source=.;Initial Catalog=\"excel to sqlDB\";Integrated Security=True;Encrypt=False;";
            string tableName = "DF"; // ← نام جدول موجود در SQL Server

            try
            {
                DataTable dt = ReadExcelToDataTable(excelPath);
                InsertToSql(dt, connectionString, tableName);

                Console.WriteLine("✅ داده‌ها با موفقیت وارد جدول شدند.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ خطا: " + ex.Message);
            }

            Console.ReadLine();
        }

        static DataTable ReadExcelToDataTable(string filePath)
        {
            DataTable dt = new DataTable();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1);
                bool firstRow = true;

                foreach (var row in worksheet.RowsUsed())
                {
                    if (firstRow)
                    {

                        foreach (var cell in row.Cells())
                            dt.Columns.Add(cell.Value.ToString());
                        firstRow = false;
                    }
                    else
                    {
                        DataRow newRow = dt.NewRow();
                        int cellCount = row.Cells().Count();

                        for (int i = 0; i < cellCount; i++)
                        {
                            var cellValue = row.Cell(i + 1).Value;
                            string text = cellValue.ToString();
                            newRow[i] = text;
                        }

                        dt.Rows.Add(newRow);
                    }
                }
            }

            return dt;
        }

        static void InsertToSql(DataTable dt, string connectionString, string tableName)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                using (SqlBulkCopy bulk = new SqlBulkCopy(conn))
                {
                    bulk.DestinationTableName = tableName;
                    bulk.BulkCopyTimeout = 0; // تا بی‌نهایت صبر می‌کنه
                    bulk.WriteToServer(dt);
                }
                conn.Close();
            }
        }
    }
}
