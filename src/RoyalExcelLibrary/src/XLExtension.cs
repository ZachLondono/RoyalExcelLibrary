using ClosedXML.Excel;
using System.Diagnostics;
using System.Windows.Forms;

namespace RoyalExcelLibrary.ExcelUI {

    public static class XLExtension {

		public static IXLCell Offset(this IXLCell cell, int rows, int columns) {
			var address = cell.Address;
			var worksheet = cell.Worksheet;
			return worksheet.Cell(address.RowNumber + rows, address.ColumnNumber + columns);
		}

		// Returns the value of the cell as a String
		public static string GetStringValue(this IXLCell cell) {

			if (cell.HasFormula) {
				return cell.CachedValue.ToString();
			}

			return cell.RichText.ToString();
			/*
						string val;
						if (cell.TryGetValue<string>(out val)) return val;
						return "";*/

		}

		/// <summary>
		/// Tries to read the value of the cell, and if it is a double that will be returned. Otherwise it will try to read the cached value, or convert the rich value into a double. 
		/// </summary>
		/// <param name="cell"></param>
		/// <returns></returns>
		public static double GetDoubleValue(this IXLCell cell) {

			try {
				object result = cell.Value;
				if (result is null) {
					Debug.WriteLine("Result is null");
				} else if (result is double) {
					return (double) result;
				}
			} catch {
				Debug.WriteLine("Could not calculate value of cell");
			}

			if (cell.HasFormula) {
				object cached = cell.CachedValue;
				if (cached is double) return (double)cached;
				return HelperFuncs.ConvertToDouble(
					cached.ToString());
			}

			double richValue = HelperFuncs.ConvertToDouble(cell.RichText.ToString());

			if (!(cell.Value is null)) {
				double value = HelperFuncs.ConvertToDouble(cell.Value.ToString());

				if (value != richValue)
					MessageBox.Show($"Unsure value for cell '{cell.Address}'. DOUBLE CHECK DIMENSIONS.", "Value Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

				return value;
			}

			return richValue;

		}

		public static string GetStringValue(this IXLWorksheet worksheet, string range) {
			var cell = worksheet.Cell(range);
			if (cell is null) return "";
			return cell.GetStringValue();
		}

	}

}
