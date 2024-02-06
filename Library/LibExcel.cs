using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework.Internal.Execution;
using OfficeOpenXml;

namespace LibraryExcel
{
    public class LibExcel
    {
        public static int[] GetIsRunData(string filePath, string sheetName)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

                List<int> rowHasRun = new List<int>();

                int rowCount = worksheet.Dimension.Rows;
                // Jika row data excel mengandung kata "Run"
                //int rowHasRun = 0;
                for (int row = 2; row <= rowCount; row++)
                {
                    if (worksheet.Cells[row, 1].Value != null &&
                        (worksheet.Cells[row, 1].Value.ToString().Trim().Equals("RUN", StringComparison.OrdinalIgnoreCase)))
                    {
                        rowHasRun.Add(row);
                    }
                }
                return rowHasRun.ToArray();
                //return rowHasRun;
            }
        }

        public static string GetDataExcel(string filePath, string columnName, string sheetName)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

                int rowCount = worksheet.Dimension.Rows;
                int columnCount = worksheet.Dimension.Columns;

                // Jika row data excel mengandung kata "Run"
                int rowHasRun = 0;
                for (int row = 2; row <= rowCount; row++)
                {
                    if (worksheet.Cells[row, 1].Value != null &&
                        (worksheet.Cells[row, 1].Value.ToString().Trim().Equals("RUN", StringComparison.OrdinalIgnoreCase)))
                    {
                        rowHasRun = row;
                        break;
                    }
                }

                // Cari data berdasarkan nama kolom
                int columnIndex = -1;
                for (int col = 1; col <= columnCount; col++)
                {
                    if (worksheet.Cells[1, col].Value.ToString() == columnName)
                    {
                        columnIndex = col;
                        break;
                    }
                }

                // Get datatable dari excel
                if (columnIndex != -1)
                {
                    return worksheet.Cells[rowHasRun, columnIndex].Value.ToString();
                }
                else
                {
                    return null; // Column tidak ditemukan
                }
            }
        }
    }
}
