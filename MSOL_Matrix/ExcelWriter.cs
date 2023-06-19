using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace MSOL_Matrix
{
    public static class ExcelWriter
    {
        public static void saveRawDtAndMatrixToExcel(string outExcelFileFolder, DataTable originalDt, DataTablesModel dataTablesModel)
        {
            string outExcelFileName = $"Cluj serials matrix ({DateTime.Now.ToString("dd MMM yyyy")}, {DateTime.Now.ToShortTimeString().Replace(":", "-")}).xlsx";

            string outExcelFullName = Path.Combine(outExcelFileFolder, outExcelFileName);

            FileInfo excelFileInfo = new FileInfo(outExcelFullName);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage pck = new ExcelPackage(excelFileInfo))
            {
                saveRawDataPage(originalDt, pck);
                saveDtToExcelAsAutoPivot(pck);
                saveDtToExcelAsMatrix(dataTablesModel.MatrixDt, pck, "Matrix_Full");

                saveDtToExcelAsMatrix(dataTablesModel.MatrixDt0Only, pck, "Matrix_0_only");
                saveDtToExcelAsMatrix(dataTablesModel.MatrixDt25, pck, "Matrix_1_25");
                saveDtToExcelAsMatrix(dataTablesModel.MatrixDt50, pck, "Matrix_26_50");
                saveDtToExcelAsMatrix(dataTablesModel.MatrixDt75, pck, "Matrix_51_75");
                saveDtToExcelAsMatrix(dataTablesModel.MatrixDt100, pck, "Matrix_76_99");
                saveDtToExcelAsMatrix(dataTablesModel.MatrixDt100Only, pck, "Matrix_100_only");

                pck.Workbook.Worksheets.MoveToStart(2);
                pck.Workbook.Worksheets.MoveToEnd(1);

                pck.Save();
            }
        }

        private static void saveRawDataPage(DataTable dt, ExcelPackage pck)
        {
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Raw data");
            ws.TabColor = System.Drawing.Color.Black;

            ws.Cells["A1"].LoadFromDataTable(dt, true);

            var start = ws.Dimension.Start;
            var end = ws.Dimension.End;

            using (ExcelRange Rng = ws.Cells[start.Row, start.Column, end.Row, end.Column])
            {
                ExcelTableCollection tblcollection = ws.Tables;
                ExcelTable table = tblcollection.Add(Rng, "tblRawData");
                table.TableStyle = TableStyles.Medium1;
                table.ShowFilter = true;
                Rng.AutoFitColumns();
            }

        }

        public static void saveDtToExcelAsAutoPivot(ExcelPackage pck)
        {
            ExcelWorksheet wsPivot;
            wsPivot = pck.Workbook.Worksheets.Add("Pivot");
            wsPivot.TabColor = System.Drawing.Color.Blue;

            var dataRange = pck.Workbook.Worksheets["Raw data"].Tables["tblRawData"];

            var pivotTable = wsPivot.PivotTables.Add(wsPivot.Cells["A1"], dataRange, "Pivot_MatrixCluj");
            pivotTable.ShowHeaders = true;
            pivotTable.UseAutoFormatting = true;
            pivotTable.ApplyWidthHeightFormats = true;
            pivotTable.ShowDrill = true;
            pivotTable.FirstHeaderRow = 1;  // first row has headers
            pivotTable.FirstDataCol = 1;    // first col of data
            pivotTable.FirstDataRow = 2;    // first row of data

            var qtyField = pivotTable.Fields["Quantity"];
            pivotTable.DataFields.Add(qtyField);

            var orderField = pivotTable.Fields["Master-order"];
            pivotTable.RowFields.Add(orderField);

            var referenceField = pivotTable.Fields["Material reference"];
            pivotTable.ColumnFields.Add(referenceField);

        }

        private static void saveDtToExcelAsMatrix(DataTable dt, ExcelPackage pck, string SheetName)
        {
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add(SheetName);
            ws.TabColor = System.Drawing.Color.LightGreen;
            ws.View.FreezePanes(2, 2);

            ws.Cells["A1"].LoadFromDataTable(dt, true);

            var start = ws.Dimension.Start;
            var end = ws.Dimension.End;

            using (ExcelRange Rng = ws.Cells[start.Row, start.Column, end.Row, end.Column])
            {
                ExcelTableCollection tblcollection = ws.Tables;
                ExcelTable table = tblcollection.Add(Rng, $"tbl{SheetName}");
                table.TableStyle = TableStyles.Medium23;
                table.ShowFilter = false;

                Rng.AutoFitColumns();

                var query = from cell in Rng
                            where cell.Value?.ToString().Contains("-1") == true
                            select cell;

                foreach (var cell in query)
                {
                    cell.Value = cell.Value.ToString().Replace("-1", "");
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                }
            }
        }
    }
}
