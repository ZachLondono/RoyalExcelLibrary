using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace RoyalExcelLibrary.ExcelUI {

	public static class ManufacturingConstants {

		public const double DadoDepth = 6;
		
		public const double SideAdj = 16;

        public const double FrontBackAdj = 0.5;

        public const double BottomAdj = 1;

        public const double SideThickness = 16;

        public const double SideSqrFtWeight = 2.1;

        public const double BottomSqrFtWeight1_4 = 0.65;

        public const double BottomSqrFtWeight1_2 = 1.55;

    }

	public static class HelperFuncs {


        // <summary>Converts a string into a double</summary>
        // <remark>
        // Attempts to use the Convert.ToDouble method, however if the string is a fraction it will do the conversion by splitting the number up into it's whole number, numerator and denominator sections and converting each to a double
        // </remark>
        public static double ConvertToDouble(string text) {

            try {
                return Convert.ToDouble(text);
            } catch (FormatException) {

                string[] parts = text.Split(' ', '/');

                double val = Convert.ToDouble(parts[0]);
                if (parts.Length == 3) {

                    double numerator = Convert.ToDouble(parts[1]);
                    double denomenator = Convert.ToDouble(parts[2]);

                    val += numerator / denomenator;

                }

                return val;

            }

        }

        /// <summary>
        /// Converts a millimeter double into fractional inches
        /// </summary>
        /// <param name="metricDim"></param>
        /// <returns></returns>
        public static string FractionalImperialDim(double metricDim) {

            // Convert to inches and round to nearest 16th
            double inches = Math.Round(metricDim / 25.4 * Math.Pow(0.0625, -1), 0) * 0.0625;

            string asString = inches.ToString();

            // If values is a whole number, just return it
            if (inches % 1 == 0) return asString;

            string[] parts = asString.Split('.');

            string x = parts[parts.Length - 1];
            if (x.Length > 5) x = x.Substring(0, 5);
            string y = "1";
            for (int i = 0; i < x.Length; i++)
                y += "0";

            int gcf = GCF(int.Parse(x), int.Parse(y));

            int numerator = int.Parse(x) / gcf;
            int denomanator = int.Parse(y) / gcf;

            if (parts[0].Equals("0")) return $"{numerator}/{denomanator}";
            return $"{parts[0]} {numerator}/{denomanator}";

        }
        private static int GCF(int x, int y) {
            x = Math.Abs(x);
            y = Math.Abs(y);
            int z;
            do {
                z = x % y;
                if (z == 0)
                    return y;
                x = y;
                y = z;
            } while (true);
        }

        public static Worksheet LoadTemplate(string path, string worksheetname, Workbook workbook) {
            try {
                Worksheet outputsheet = workbook.Worksheets[worksheetname];
                outputsheet.Delete();
            } catch (COMException) {
                Debug.WriteLine("Output sheet could be deleted or does not exist");
            }
            Application app = (Application)ExcelDnaUtil.Application;
            Workbook template = app.Workbooks.Open(path);
            template.Worksheets[worksheetname].Copy(workbook.Worksheets[workbook.Worksheets.Count - 1]);
            template.Close();
            return workbook.Worksheets[worksheetname];
        }

    }

}
