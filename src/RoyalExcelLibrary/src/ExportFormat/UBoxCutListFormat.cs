using RoyalExcelLibrary.ExcelUI.Models;
using RoyalExcelLibrary.ExcelUI.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Microsoft.Office.Interop.Excel;

namespace RoyalExcelLibrary.ExcelUI.ExportFormat {
    public class UBoxCutListFormat : StdCutListFormat {

        public override Range WriteOrderParts(IEnumerable<string[,]> seperatedBoxes, Worksheet outputsheet, int startRow, int startCol) {

            int currRow = startRow;

            // Format the header
            string[] box_headers = new string[] { "cab#", "Part Name", "Comment", "Qty", "Width", "Length", "Material", "Line#", "Top Down Diagram" };
            Range rng = outputsheet.Range[outputsheet.Cells[currRow, startCol], outputsheet.Cells[currRow++, startCol + box_headers.Length - 1]];
            rng.Value = box_headers;
            rng.Interior.Color = Highlightcolor;
            rng.EntireRow.RowHeight = 35;
            rng.VerticalAlignment = XlVAlign.xlVAlignCenter;

            int i = 1;
            foreach (string[,] boxRows in seperatedBoxes) {
                int rows = boxRows.GetLength(0);
                int cols = boxRows.GetLength(1);

                rng = outputsheet.Range[outputsheet.Cells[currRow, startCol], outputsheet.Cells[currRow + rows - 1, startCol + cols - 1]];
                rng.Value = boxRows;
                if (i++ % 2 == 0) rng.Interior.Color = Highlightcolor;

                double A = Convert.ToDouble(rng.Offset[1][6].Value2);
                double B = Convert.ToDouble(rng.Offset[3][6].Value2);
                double C = Convert.ToDouble(rng.Offset[5][6].Value2);
                try {
                    var diagram = AddUBoxDiagram(A, B, C, outputsheet);
                    diagram.Left = (float)outputsheet.Range["I1"].Left;
                    diagram.Top = (float)rng.Top;
                    diagram.Height = (float) rng.Height;
                    diagram.Width = (float)outputsheet.Range["I1"].Width;
                } catch {
                    Debug.WriteLine("Unable to add U-Box Diagram. Check that the image file is still accessable");
				}

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

            // Increase size of qty and dimensions
            Range dimRng = outputsheet.Range[outputsheet.Cells[startRow, startCol + 3], outputsheet.Cells[startRow, startCol + 3]];
            for (int o = 0; o < 3; o++) {
                dimRng.Offset[0, o].EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                if (dimRng.Offset[0, o].ColumnWidth < 6) dimRng.Offset[0, o].ColumnWidth = 6;
            }

            // Increase the size of the UBox image
            var diagramRng = outputsheet.Range[outputsheet.Cells[startRow, startCol + 8], outputsheet.Cells[startRow, startCol + 8]];
            if (diagramRng.ColumnWidth < 25) diagramRng.ColumnWidth = 25;

            return fullRng;

        }

        private Shape AddUBoxDiagram(double A, double B, double C, Worksheet sheet) {

            string frac_A = HelperFuncs.FractionalImperialDim(A);
            string frac_B = HelperFuncs.FractionalImperialDim(B);
            string frac_C = HelperFuncs.FractionalImperialDim(C);

            var image = sheet.Shapes.AddPicture("R:\\DB ORDERS\\Images\\BlankUbox.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 100, 100);

            Shape CreateTextBox(string value) {
                var textbox = sheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 10, 10);
                textbox.TextFrame.Characters(Type.Missing, Type.Missing).Text = value;
                textbox.TextFrame2.TextRange.Font.Size = 8;
                textbox.TextFrame2.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                textbox.TextFrame.AutoSize = true;
                textbox.Fill.Transparency = 1;
                textbox.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                return textbox;
            }

            var textbox_A = CreateTextBox(frac_A);
            textbox_A.Left = 5;
            textbox_A.Top = 0;
            var textbox_B = CreateTextBox(frac_B);
            textbox_B.Left = image.Width - textbox_B.Width - 5;
            textbox_B.Top = 0;
            var textbox_C = CreateTextBox(frac_C);
            textbox_C.Left = image.Width / 2 - textbox_C.Width / 2;
            textbox_C.Top = image.Height / 2 - 15;

            var shapes = new string[] { textbox_A.Name, textbox_B.Name, textbox_C.Name, image.Name };
            ShapeRange shapeRange = sheet.Shapes.Range[shapes];
            var group = shapeRange.Group();

            return group;

        }


    }

}
