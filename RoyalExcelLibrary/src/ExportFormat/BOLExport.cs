using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using RoyalExcelLibrary.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.ExportFormat {
	public class BOLExport : IExcelExport {

		public readonly string _bolTemplateFile = "R:\\DB ORDERS\\RoyalExcelLibrary\\Export Templates\\\\BOLTemplate.xlsx";

		public Worksheet ExportOrder(Order order, Workbook workbook) {

			Worksheet outputsheet;
			string worksheetname = "BOL";

			outputsheet = HelperFuncs.LoadTemplate(_bolTemplateFile, worksheetname, workbook);

			FillField(outputsheet.Range["Consignee"], "TO CONSIGNEE", order.Customer.Name);
			FillField(outputsheet.Range["Address1"], "STREET", order.Customer.Address.Line1);
			FillField(outputsheet.Range["Address2"], "STREET", order.Customer.Address.Line2);
			FillField(outputsheet.Range["CityState"], "DESTINATION: CITY & STATE", order.Customer.Address.City + ", " + order.Customer.Address.State);
			FillField(outputsheet.Range["Zip"], "ZIP CODE", order.Customer.Address.Zip);
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

			var contentChars = range.Characters[header.Length + 1];
			contentChars.Font.Name = "Arial";
			contentChars.Font.FontStyle = "Regular";
			contentChars.Font.Size = 10;
			contentChars.Font.Bold = false;

		}

	}
}
