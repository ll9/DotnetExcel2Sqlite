using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinformExcelSqliteExample
{
    class ExcelHandler
    {
        public static DataTable ImportExceltoDatatable(string filePath)
        {
            // Open the Excel file using ClosedXML.
            // Keep in mind the Excel file cannot be open when trying to read it
            using (XLWorkbook workBook = new XLWorkbook(filePath))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(1);

                //Create a new DataTable.
                DataTable dataTable = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dataTable.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        //Add rows to DataTable.
                        dataTable.Rows.Add();
                        int i = 0;

                        foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                        {
                            dataTable.Rows[dataTable.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }
                }

                return dataTable;
            }
        }

        public static void ExportExcelToSqlite(string excelPath, string sqlitePath, string tableName)
        {
            using (var connection = GetConnection(sqlitePath))
            {
                connection.Open();
                using (XLWorkbook workBook = new XLWorkbook(excelPath))
                {
                    IXLWorksheet workSheet = workBook.Worksheet(1);

                    var firstRow = true;
                    foreach (IXLRow row in workSheet.Rows())
                    {
                        //Use the first row to add Headings to Sqlite
                        if (firstRow)
                        {
                            var query = String.Format("CREATE TABLE {0} ({1});",
                                tableName,
                                String.Join(", ",
                                row.Cells().Select(cell => $"{cell.Value.ToString()} VARCHAR(30)")
                                )
                            );
                            new SQLiteCommand(query, connection).ExecuteNonQuery();
                            firstRow = false;
                        }
                        else
                        // Insert Values into table
                        {
                            var query = String.Format("INSERT INTO {0} VALUES ({1});",
                                tableName,
                                String.Join(", ",
                                row.Cells().Select(cell => $"'{cell.Value.ToString()}'")
                                ));

                            new SQLiteCommand(query, connection).ExecuteNonQuery();
                        }
                    }
                }

            }
        }

        private static SQLiteConnection GetConnection(string sqlitePath)
        {
            return new SQLiteConnection($"Data Source={sqlitePath}");
        }
    }
}
