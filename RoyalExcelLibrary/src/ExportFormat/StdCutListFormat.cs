using Microsoft.Office.Interop.Excel;
using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System.Collections.Generic;
using System.Linq;

namespace RoyalExcelLibrary.ExportFormat {
	public class StdCutListFormat : ICutListFormat {

        public System.Drawing.Color Highlightcolor { get; set; }
        public StdCutListFormat() {
            Highlightcolor = System.Drawing.Color.FromArgb(191,191,191);
		}

        public Range WriteOrderHeader(Order order, Worksheet outputsheet) {

            Range rng = outputsheet.Range["B1"];
            rng.Value = "Company";
            rng.Interior.Color = Highlightcolor;

            rng = outputsheet.Range["C1:G1"];
            rng.Merge();
            rng.Value = order.CustomerName;

            rng = outputsheet.Range["B2", "B3"];
            rng.Value = new string[,] { { "Order#" }, { "Job Name" } };
            rng.Interior.Color = Highlightcolor;

            var ordernum = outputsheet.Range["C2", "D2"];
            var jobname = outputsheet.Range["C3", "D3"];
            ordernum.Merge();
            ordernum.Value = order.Number;
            jobname.Merge();
            jobname.Value = order.Job.Name;

            rng = outputsheet.Range["E2", "E3"];
            rng.Value = new string[,] { { "Date" }, { "Box Count" } };
            rng.Interior.Color = Highlightcolor;

            var date = outputsheet.Range["F2", "G2"];
            var boxcount = outputsheet.Range["F3", "G3"];
            date.Merge();
            date.Value = order.Job.CreationDate;
            date.NumberFormat = "mm/dd/yy";
            boxcount.Merge();
            boxcount.Value = order.Products.Where(p => p is DrawerBox)
                                            .Select(b => (b as DrawerBox).Qty)
                                            .Sum();

            rng = outputsheet.Range["B1", "G3"];
            rng.Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            rng.RowHeight = 35;
            rng.VerticalAlignment = XlVAlign.xlVAlignCenter;
            rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rng.WrapText = true;

            return rng;

        }

        public virtual Range WriteOrderParts(IEnumerable<string[,]> seperatedBoxes, Worksheet outputsheet, int startRow, int startCol) {
            string[] box_headers = new string[] { "cab#", "Part Name", "Comment", "Qty", "Width", "Length", "Material", "Line#", "Box Size" };

            int currRow = startRow;
            Range rng = outputsheet.Range[outputsheet.Cells[currRow, startCol], outputsheet.Cells[currRow++, startCol + box_headers.Length - 1]];
            rng.Value = box_headers;
            rng.Interior.Color = Highlightcolor;
            rng.EntireRow.RowHeight = 35;
            rng.VerticalAlignment = XlVAlign.xlVAlignCenter;
            rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rng.Cells.Borders.LineStyle = XlLineStyle.xlContinuous;

            int i = 1;
            foreach (string[,] boxRows in seperatedBoxes) {
                int rows = boxRows.GetLength(0);
                int cols = boxRows.GetLength(1);

                rng = outputsheet.Range[outputsheet.Cells[currRow, startCol], outputsheet.Cells[currRow + rows - 1, startCol + cols - 1]];
                rng.Value = boxRows;
                rng.Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
                rng.Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                if (i++ % 2 == 0) rng.Interior.Color = Highlightcolor;

                currRow += rows;
            }

            // Auto fit part header columns
            Range fullRng = outputsheet.Range[outputsheet.Cells[startRow, startCol], outputsheet.Cells[currRow, startCol + 8]];
            fullRng.Columns.AutoFit();

            // Auto fit each part row
            Range partRng = outputsheet.Range[outputsheet.Cells[startRow + 1, startCol], outputsheet.Cells[currRow, startCol + 8]];
            partRng.Rows.AutoFit();

            // Make sure comment column has space for writing extra comments
            Range comRng = outputsheet.Range[outputsheet.Cells[startRow, startCol + 2], outputsheet.Cells[startRow, startCol + 2]];
            if (comRng.ColumnWidth < 30) comRng.ColumnWidth = 30;

            // Increase size of qty and dimensions
            Range dimRng = outputsheet.Range[outputsheet.Cells[startRow, startCol + 3], outputsheet.Cells[startRow, startCol + 3]];
            for (int o = 0; o < 3; o++) {
                dimRng.Offset[0, o].EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                if (dimRng.Offset[0, o].ColumnWidth < 10) dimRng.Offset[0, o].ColumnWidth = 10;
            }

            // Increase the size of the box size column
            Range sizeRng = outputsheet.Range[outputsheet.Cells[startRow, startCol + 8], outputsheet.Cells[startRow, startCol + 8]];
            sizeRng.Columns.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            if (sizeRng.ColumnWidth < 25) sizeRng.ColumnWidth = 25;

            return fullRng;
        }
    }

}
