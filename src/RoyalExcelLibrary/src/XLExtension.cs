using ClosedXML.Excel;
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

		public static double GetDoubleValue(this IXLCell cell) {

			if (cell.HasFormula) {
				return HelperFuncs.ConvertToDouble(
					cell.CachedValue.ToString());
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
