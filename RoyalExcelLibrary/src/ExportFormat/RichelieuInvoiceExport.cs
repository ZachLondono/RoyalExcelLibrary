using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace RoyalExcelLibrary.ExportFormat {
	public class RichelieuInvoiceExport : IExcelExport {

		public readonly string _invoiceTemplateFile = "R:\\DB ORDERS\\RoyalExcelLibrary\\Export Templates\\RichelieuInvoiceTemplate.xlsx";

		public Worksheet ExportOrder(Order order, Workbook workbook) {

			Worksheet outputsheet;
			string worksheetname = "Invoice";

			outputsheet = HelperFuncs.LoadTemplate(_invoiceTemplateFile, worksheetname, workbook);

			outputsheet.Range["RefNum"].Value2 = order.Job.Name;
			outputsheet.Range["InvoiceNum"].Value2 = (order as RichelieuOrder).RichelieuNumber;
			outputsheet.Range["PONum"].Value2 = order.Number;

			IEnumerable<DrawerBox> boxes = order.Products.Where(p => p is DrawerBox).Cast<DrawerBox>();

			Range skuStart = outputsheet.Range["SkuStart"];
			Range descStart = outputsheet.Range["DescriptionStart"];
			Range qtyStart = outputsheet.Range["QtyStart"];
			Range heightStart = outputsheet.Range["HeightStart"];
			Range widthStart = outputsheet.Range["WidthStart"];
			Range depthStart = outputsheet.Range["DepthStart"];
			Range priceStart = outputsheet.Range["PriceStart"];

			int i = 0;
			foreach (DrawerBox box in boxes) {
				skuStart.Offset[i, 0].Value2 = box.ProductName;
				descStart.Offset[i, 0].Value2 = box.ProductDescription?.Replace('\n', ',') ?? "";
				qtyStart.Offset[i, 0].Value2 = box.Qty;
				heightStart.Offset[i, 0].Value2 = box.Height / 25.4;
				widthStart.Offset[i, 0].Value2 = box.Width / 25.4;
				depthStart.Offset[i, 0].Value2 = box.Depth / 25.4;
				priceStart.Offset[i, 0].Value2 = box.UnitPrice;
				i++;
			}

			Range print_rng = outputsheet.Range[outputsheet.Cells[1, 1], outputsheet.Cells[i+skuStart.Row, priceStart.Column + 2]];
			outputsheet.PageSetup.PrintArea = print_rng.Address;

			return outputsheet;

		}

	}

}
