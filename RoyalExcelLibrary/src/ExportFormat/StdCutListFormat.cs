using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System.Collections.Generic;
using System.Linq;

using Excel = Microsoft.Office.Interop.Excel;

namespace RoyalExcelLibrary.ExportFormat {
	public class StdCutListFormat : ICutListFormat {

        public System.Drawing.Color Highlightcolor { get; set; }
        public StdCutListFormat() {
            Highlightcolor = System.Drawing.Color.FromArgb(191,191,191);
		}

        public Excel.Range WriteOrderHeader(Order order, Excel.Worksheet outputsheet) {

            Excel.Range rng = outputsheet.Range["B1"];
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
            rng.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            rng.RowHeight = 35;
            rng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            rng.WrapText = true;

            return rng;

        }

        public virtual Excel.Range WriteOrderParts(IEnumerable<string[,]> seperatedBoxes, Excel.Worksheet outputsheet, int startRow, int startCol) {
            string[] box_headers = new string[] { "cab#", "Part Name", "Comment", "Qty", "Width", "Length", "Material", "Line#", "Box Size" };

            int currRow = startRow;
            Excel.Range rng = outputsheet.Range[outputsheet.Cells[currRow, startCol], outputsheet.Cells[currRow++, startCol + box_headers.Length - 1]];
            rng.Value = box_headers;
            rng.Interior.Color = Highlightcolor;
            rng.EntireRow.RowHeight = 35;
            rng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            rng.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            int i = 1;
            foreach (string[,] boxRows in seperatedBoxes) {
                int rows = boxRows.GetLength(0);
                int cols = boxRows.GetLength(1);

                rng = outputsheet.Range[outputsheet.Cells[currRow, startCol], outputsheet.Cells[currRow + rows - 1, startCol + cols - 1]];
                rng.Value = boxRows;
                rng.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                rng.Cells.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                if (i++ % 2 == 0) rng.Interior.Color = Highlightcolor;

                currRow += rows;
            }

            // Auto fit part header columns
            var fullRng = outputsheet.Range[outputsheet.Cells[startRow, startCol], outputsheet.Cells[currRow, startCol + 8]];
            fullRng.Columns.AutoFit();

            // Auto fit each part row
            var partRng = outputsheet.Range[outputsheet.Cells[startRow + 1, startCol], outputsheet.Cells[currRow, startCol + 8]];
            partRng.Rows.AutoFit();

            // Make sure comment column has space for writing extra comments
            var comRng = outputsheet.Range[outputsheet.Cells[startRow, startCol + 2], outputsheet.Cells[startRow, startCol + 2]];
            if (comRng.ColumnWidth < 30) comRng.ColumnWidth = 30;

            // Increase the size of the box size column
            Excel.Range sizeRng = outputsheet.Range[outputsheet.Cells[startRow, startCol + 8], outputsheet.Cells[startRow, startCol + 8]];
            sizeRng.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            if (sizeRng.ColumnWidth < 25) sizeRng.ColumnWidth = 25;

            return fullRng;
        }
    }

}
