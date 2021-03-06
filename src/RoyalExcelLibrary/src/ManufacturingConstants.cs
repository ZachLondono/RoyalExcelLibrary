using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace RoyalExcelLibrary.ExcelUI {

	public static class HelperFuncs {

        public static AppSettings ReadSettings() {

            string settingsPath = Path.Combine(Path.GetDirectoryName(ExcelDnaUtil.XllPath), "appsettings.json");

            using (var reader = new StreamReader(settingsPath)) {
                return JsonConvert.DeserializeObject<AppSettings>(reader.ReadToEnd());
            }
        }

        // <summary>Converts a string into a double</summary>
        // <remark>
        // Attempts to use the Convert.ToDouble method, however if the string is a fraction it will do the conversion by splitting the number up into it's whole number, numerator and denominator sections and converting each to a double
        // </remark>
        public static double ConvertToDouble(string text) {

            try {
                return Convert.ToDouble(text);
            } catch (FormatException) {

                // If the text number has double spaces or a leading/trailing space it will not be parsed correctly
                string fixedStr = text.Trim().Replace("  ", " ");

                string[] parts = fixedStr.Split(' ', '/');

                double val = Convert.ToDouble(parts[0]);
                if (parts.Length == 3) {

                    double numerator = Convert.ToDouble(parts[1]);
                    double denomenator = Convert.ToDouble(parts[2]);

                    val += numerator / denomenator;

                } else {

                    MessageBox.Show($"Error parsing number value '{text}'. DOUBLE CHECK DIMENSIONS", "Dimension Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

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

            double accuracy = (1.0 / 32.0);

            // Convert to inches and round to nearest 16th
            double inches = Math.Round(metricDim / 25.4 * Math.Pow(accuracy, -1), 0) * accuracy;

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
            Microsoft.Office.Interop.Excel.Application app = (Microsoft.Office.Interop.Excel.Application)ExcelDnaUtil.Application;
            Workbook template = app.Workbooks.Open(path);
            template.Worksheets[worksheetname].Copy(workbook.Worksheets[workbook.Worksheets.Count - 1]);
            template.Close();
            return workbook.Worksheets[worksheetname];
        }

    }

}
