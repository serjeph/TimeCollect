using ClosedXML.Excel;
using System;
using System.Collections.Generic;

namespace TimeCollect.Helpers
{
    internal class ExcelHelper
    {
        public static void ExportToExcel(IList<IList<object>> data, string sheetName, string filepath, List<string> columnHeaders = null)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add(sheetName);
                    int row = 1;

                    // Add column headers
                    if (columnHeaders != null)
                    {
                        int col = 1;
                        foreach (var header in columnHeaders)
                        {
                            worksheet.Cell(row, col).Value = header;
                            col++;
                        }
                        row++;
                    }

                    foreach (var rowData in data)
                    {
                        int col = 1;
                        foreach (var cellData in rowData)
                        {
                            switch (cellData)
                            {
                                case string str:
                                    worksheet.Cell(row, col).Value = str;
                                    break;
                                case int num:
                                    worksheet.Cell(row, col).Value = num;
                                    break;
                                case float f:
                                    worksheet.Cell(row, col).Value = Math.Round(f, 2);
                                    break;
                                default:
                                    worksheet.Cell(row, col).SetValue(cellData.ToString());
                                    break;
                            }
                            col++;
                        }
                        row++;
                    }

                    workbook.SaveAs(filepath);
                }
            }
            catch (Exception ex)
            {
                // Handle exceptions
                Console.WriteLine($"Error exporting to Excel: {ex.Message}");
            }
        }
    }
}
