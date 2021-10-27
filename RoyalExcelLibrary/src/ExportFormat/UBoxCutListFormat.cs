﻿using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

using Excel = Microsoft.Office.Interop.Excel;

namespace RoyalExcelLibrary.ExportFormat {
    public class UBoxCutListFormat : StdCutListFormat {

        public override Excel.Range WriteOrderParts(IEnumerable<string[,]> seperatedBoxes, Excel.Worksheet outputsheet, int startRow, int startCol) {
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

            return outputsheet.Range[outputsheet.Cells[startRow, startCol], outputsheet.Cells[currRow, startCol + 8]];
        }

        private Excel.Shape AddUBoxDiagram(double A, double B, double C, Excel.Worksheet sheet) {

            string frac_A = HelperFuncs.FractionalImperialDim(A);
            string frac_B = HelperFuncs.FractionalImperialDim(B);
            string frac_C = HelperFuncs.FractionalImperialDim(C);

            var image = sheet.Shapes.AddPicture("R:\\DB ORDERS\\Images\\BlankUbox.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 0, 0, 100, 100);

            Excel.Shape CreateTextBox(string value) {
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
            Excel.ShapeRange shapeRange = sheet.Shapes.Range[shapes];
            var group = shapeRange.Group();

            return group;

        }


    }

}
