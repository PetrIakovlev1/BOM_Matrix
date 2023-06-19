using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace MSOL_Matrix
{
    public static class ExcelReader
    {
        public static DataTable tryReadExcel(FileInfo excelFileInfo, string readResult)
        {
            DataTable dt = new DataTable();
            List<string> columnToReadList = new List<string>() { "MASTER-ORDER", "MATERIAL REFERENCE", "QUANTITY" };
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var pck = new ExcelPackage(excelFileInfo))
                {
                    var ws = pck.Workbook.Worksheets.First();

                    foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    {
                        if (firstRowCell.Text.ToUpper().Trim() == "QUANTITY")
                        {
                            dt.Columns.Add(firstRowCell.Text, typeof(int));
                        }
                        else
                        {
                            dt.Columns.Add(firstRowCell.Text);
                        }
                    }

                    var startRow = 2;
                    for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                        DataRow row = dt.Rows.Add();
                        foreach (var cell in wsRow)
                        {
                            row[cell.Start.Column - 1] = cell.Text;
                        }
                    }
                }

                List<string> columnsToRemove = new List<string>();
                foreach (DataColumn dc in dt.Columns)
                {
                    if (columnToReadList.IndexOf(dc.ColumnName.ToString().ToUpper().Trim()) < 0)
                        columnsToRemove.Add(dc.ColumnName);
                }

                foreach (string c in columnsToRemove)
                {
                    dt.Columns.Remove(c);
                }

                List<int> RowsIndexToDelList = new List<int>();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.Rows[i].ItemArray.Any(x => x is DBNull || string.IsNullOrWhiteSpace(x?.ToString())))
                        RowsIndexToDelList.Add(i);
                }

                for (int i = RowsIndexToDelList.Count - 1; i >= 0; i--)
                {
                    dt.Rows.RemoveAt(RowsIndexToDelList[i]);
                }

                dt.AcceptChanges();
            }
            catch (Exception ex)
            {
                readResult = ex.Message;
            }
            return dt;
        }
    }
}
