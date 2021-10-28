using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System.Collections.Generic;
using System.Linq;

using Excel = Microsoft.Office.Interop.Excel;

namespace RoyalExcelLibrary.ExportFormat {
	public class StdCutListFormat : ICutListFormat {

        public System.Drawing.Color Highlightcolor { get; set; }
        public StdCutListFormat() {
            Highlightcolor = System.Drawing.Color.LightGray;
		}

        public Excel.Range WriteOrderHeader(Order order, Excel.Worksheet outputsheet) {

            Excel.Range rng = outputsheet.Range["B1", "B2"];
            rng.Value = new string[,] { { "Order#" }, { "Job Name" } };
            rng.Interior.Color = Highlightcolor;

            var ordernum = outputsheet.Range["C1", "D1"];
            var jobname = outputsheet.Range["C2", "D2"];
            ordernum.Merge();
            ordernum.Value = order.Number;
            jobname.Merge();
            jobname.Value = order.Job.Name;

            rng = outputsheet.Range["E1", "E2"];
            rng.Value = new string[,] { { "Date" }, { "Box Count" } };
            rng.Interior.Color = Highlightcolor;

            var date = outputsheet.Range["F1", "G1"];
            var boxcount = outputsheet.Range["F2", "G2"];
            date.Merge();
            date.Value = order.Job.CreationDate;
            date.NumberFormat = "mm/dd/yy";
            boxcount.Merge();
            boxcount.Value = order.Products.Where(p => p is DrawerBox)
                                            .Select(b => (b as DrawerBox).Qty)
                                            .Sum();

            rng = outputsheet.Range["B1", "G2"];
            rng.Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            rng.RowHeight = 35;
            rng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            rng.WrapText = true;

            return rng;

        }

        public virtual Excel.Range WriteOrderParts(IEnumerable<string[,]> seperatedBoxes, Excel.Worksheet outputsheet, int startRow, int startCol) {
            string[] box_headers = new string[] { "cab#", "part", "comment", "qty", "width", "length", "material", "line#", "box size" };

            int currRow = startRow;
            Excel.Range rng = outputsheet.Range[outputsheet.Cells[currRow, startCol], outputsheet.Cells[currRow++, startCol + box_headers.Length - 1]];
            rng.Value = box_headers;
            rng.Interior.Color = Highlightcolor;
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

            return outputsheet.Range[outputsheet.Cells[startRow, startCol], outputsheet.Cells[currRow, startCol + 8]];
        }
    }

}
