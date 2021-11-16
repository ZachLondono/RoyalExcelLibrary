using Microsoft.Office.Interop.Excel;
using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Options;
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

            rng = outputsheet.Range["C1:F1"];
            rng.Merge();
            rng.Value = order.Customer.Name;

            rng = outputsheet.Range["G1"];
            rng.Value = "Vendor";
            rng.Interior.Color = Highlightcolor;

            rng = outputsheet.Range["H1:I1"];
            rng.Merge();
            rng.Value = order.Vendor?.Name ?? "";

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

            IEnumerable<DrawerBox> boxes = order.Products.Where(p => p is DrawerBox).Cast<DrawerBox>();

            var date = outputsheet.Range["F2", "G2"];
            var boxcount = outputsheet.Range["F3", "G3"];
            date.Merge();
            date.Value = order.Job.CreationDate;
            date.NumberFormat = "mm/dd/yy";
            boxcount.Merge();
            boxcount.Value = boxes.Select(b => b.Qty)
                                    .Sum();

            UndermountNotch mostCommonUM = boxes.GroupBy(b => b.NotchOption)
                                            .OrderByDescending(bg => bg.Count())
                                            .Select(bg => bg.Key)
                                            .FirstOrDefault();

            Clips mostCommonClip = boxes.GroupBy(b => b.ClipsOption)
                                            .OrderByDescending(bg => bg.Count())
                                            .Select(bg => bg.Key)
                                            .FirstOrDefault();

            bool mostCommonHoles = boxes.GroupBy(b => b.MountingHoles)
                                            .OrderByDescending(bg => bg.Count())
                                            .Select(bg => bg.Key)
                                            .FirstOrDefault();

            bool mostCommonFinish = boxes.GroupBy(b => b.PostFinish)
                                            .OrderByDescending(bg => bg.Count())
                                            .Select(bg => bg.Key)
                                            .FirstOrDefault();

            rng = outputsheet.Range["H2:I2"];
            rng.Merge();
            rng.Value2 = mostCommonUM.ToString() ?? "";

            rng = outputsheet.Range["H3:I3"];
            rng.Merge();
            rng.Value2 = $"clips:{mostCommonClip.ToString() ?? ""}";

            rng = outputsheet.Range["H4:I4"];
            rng.Merge();
            rng.Value2 = $"Mounting Holes: {(mostCommonHoles ? "Yes" : "No" )}";

            rng = outputsheet.Range["B4"];
            rng.Interior.Color = Highlightcolor;
            rng.Value2 = "Note";

            rng = outputsheet.Range["C4:E4"];
            rng.Merge();
            rng.Value2 = "";//order.Note;

            rng = outputsheet.Range["F4:G4"];
            rng.Merge();
            rng.Value2 = $"Post Finish: {(mostCommonFinish ? "Yes" : "No")}";

            rng = outputsheet.Range["B1", "I4"];
            rng.RowHeight = 35;
            rng.VerticalAlignment = XlVAlign.xlVAlignCenter;
            rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rng.WrapText = true;
            rng.Borders.LineStyle = XlLineStyle.xlContinuous;

            return rng;

        }

        public virtual Range WriteOrderParts(IEnumerable<string[,]> seperatedBoxes, Worksheet outputsheet, int startRow, int startCol) {
            string[] box_headers = new string[] { "cab#", "Part Name", "Comment", "Qty", "Width", "Length", "Material", "Line#", "Box/Part Size" };

            int currRow = startRow;
            Range rng = outputsheet.Range[outputsheet.Cells[currRow, startCol], outputsheet.Cells[currRow++, startCol + box_headers.Length - 1]];
            rng.Value = box_headers;
            rng.Interior.Color = Highlightcolor;
            rng.EntireRow.RowHeight = 35;
            rng.VerticalAlignment = XlVAlign.xlVAlignCenter;
            rng.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            int i = 1;
            foreach (string[,] boxRows in seperatedBoxes) {
                int rows = boxRows.GetLength(0);
                int cols = boxRows.GetLength(1);

                rng = outputsheet.Range[outputsheet.Cells[currRow, startCol], outputsheet.Cells[currRow + rows - 1, startCol + cols - 1]];
                rng.Value = boxRows;
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

            // Make sure the cab num column is not too big
            Range cabNumRng = outputsheet.Range[outputsheet.Cells[startRow, startCol], outputsheet.Cells[currRow, startCol]];
            if (cabNumRng.ColumnWidth > 15) {
                cabNumRng.ColumnWidth = 15;
                cabNumRng.WrapText = true;
            }

            // Increase size of qty and dimensions
            Range dimRng = outputsheet.Range[outputsheet.Cells[startRow, startCol + 3], outputsheet.Cells[startRow, startCol + 3]];
            for (int o = 0; o < 3; o++) {
                dimRng.Offset[0, o].EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                if (dimRng.Offset[0, o].ColumnWidth < 6) dimRng.Offset[0, o].ColumnWidth = 6;
            }

            // Increase the size of the box size column
            Range sizeRng = outputsheet.Range[outputsheet.Cells[startRow, startCol + 8], outputsheet.Cells[startRow, startCol + 8]];
            sizeRng.Columns.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            if (sizeRng.ColumnWidth < 25) sizeRng.ColumnWidth = 25;

            return fullRng;
        }
    }

}
