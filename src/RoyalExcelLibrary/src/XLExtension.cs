using ClosedXML.Excel;

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

			string value;

			if (cell.HasFormula) {
				value = cell.CachedValue.ToString();
			} else {
				value = cell.RichText.ToString();
			}

			return HelperFuncs.ConvertToDouble(value);


		}

		public static string GetStringValue(this IXLWorksheet worksheet, string range) {
			var cell = worksheet.Cell(range);
			if (cell is null) return "";
			return cell.GetStringValue();
		}

	}

}
