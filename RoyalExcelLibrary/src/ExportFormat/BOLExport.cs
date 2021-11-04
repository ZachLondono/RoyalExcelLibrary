using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using RoyalExcelLibrary.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.ExportFormat {
	public class BOLExport : IExcelExport {

		public readonly string _bolTemplateFile = "R:\\DB ORDERS\\RoyalExcelLibrary\\Export Templates\\\\BOLTemplate.xlsx";

		public Worksheet ExportOrder(Order order, ExportData data, Workbook workbook) {

			Worksheet outputsheet;
			string worksheetname = "BOL";

			try {
				outputsheet = workbook.Worksheets[worksheetname];
			} catch (COMException) {
				Application app = (Application)ExcelDnaUtil.Application;
				Workbook template = app.Workbooks.Open(_bolTemplateFile);
				template.Worksheets[worksheetname].Copy(workbook.Worksheets[workbook.Worksheets.Count - 1]);
				template.Close();
				outputsheet = workbook.Worksheets[worksheetname];
			}

			FillField(outputsheet.Range["Consignee"], "TO CONSIGNEE", order.CustomerName);
			FillField(outputsheet.Range["Address1"], "STREET", order.ShipAddress.StreetAddress);
			FillField(outputsheet.Range["Address2"], "STREET", "");
			FillField(outputsheet.Range["CityState"], "DESTINATION: CITY & STATE", order.ShipAddress.City + ", " + order.ShipAddress.State);
			FillField(outputsheet.Range["Zip"], "ZIP CODE", order.ShipAddress.Zip);
			FillField(outputsheet.Range["PhoneNum"], "PHONE", "");
			FillField(outputsheet.Range["RefNum"], "REF#", "");

			return outputsheet;

		}

		private void FillField(Range range, string header, string content) {

			range.Value2 = header + "\n" + content;

			var headerChars = range.Characters[0, header.Length];
			headerChars.Font.Name = "Arial";
			headerChars.Font.FontStyle = "Regular";
			headerChars.Font.Size = 7;
			headerChars.Font.Bold = true;

			var contentChars = range.Characters[header.Length];
			contentChars.Font.Name = "Arial";
			contentChars.Font.FontStyle = "Regular";
			contentChars.Font.Size = 10;
			contentChars.Font.Bold = false;

		}

	}
}
