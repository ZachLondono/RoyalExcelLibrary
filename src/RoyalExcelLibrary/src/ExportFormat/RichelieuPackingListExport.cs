using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using RoyalExcelLibrary.ExcelUI.Models;
using RoyalExcelLibrary.ExcelUI.Models.Products;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.ExcelUI.ExportFormat {
	public class RichelieuPackingListExport : IExcelExport {


		public readonly string _packinglistTemplateFile = "R:\\DB ORDERS\\RoyalExcelLibrary\\Export Templates\\RichelieuPackingListTemplate.xlsx";

		public Worksheet ExportOrder(Order order, Workbook workbook) {

			Worksheet outputsheet;
			string worksheetname = "Packing List";

			outputsheet = HelperFuncs.LoadTemplate(_packinglistTemplateFile, worksheetname, workbook);

			RichelieuOrder richOrder = (RichelieuOrder)order;

			outputsheet.Range["Customer"].Value2 = ""; // Customer #
			outputsheet.Range["Name"].Value2 = richOrder.ClientLastName + ", " + richOrder.ClientFirstName; // First Name / Last Name
			outputsheet.Range["Company"].Value2 = order.Customer.Name;
			var addr = order.Customer.Address;
			outputsheet.Range["Address"].Value2 = addr.Line1 + " " + addr.Line2 + " " + addr.City + ", " + addr.State + " " + addr.Zip;
			outputsheet.Range["OrderNum"].Value2 = $"{(order as RichelieuOrder).RichelieuNumber} / {order.Number} / {richOrder.ClientPurchaseOrder} / ";
			
			IEnumerable<DrawerBox> boxes = order.Products.Where(p => p is DrawerBox).Cast<DrawerBox>();

			Range skuStart = outputsheet.Range["SkuStart"];
			Range descStart = outputsheet.Range["DescriptionStart"];
			Range qtyStart = outputsheet.Range["QtyStart"];
			Range heightStart = outputsheet.Range["HeightStart"];
			Range widthStart = outputsheet.Range["WidthStart"];
			Range depthStart = outputsheet.Range["DepthStart"];

			int i = 0;
			foreach (DrawerBox box in boxes) {
				skuStart.Offset[i, 0].Value2 = box.ProductName;
				descStart.Offset[i, 0].Value2 = box.ProductDescription?.Replace('\n', ',') ?? "";
				qtyStart.Offset[i, 0].Value2 = box.Qty;
				heightStart.Offset[i, 0].Value2 = box.Height / 25.4;
				widthStart.Offset[i, 0].Value2 = box.Width / 25.4;
				depthStart.Offset[i, 0].Value2 = box.Depth / 25.4;
				i++;
			}

			Range print_rng = outputsheet.Range[outputsheet.Cells[1, 1], outputsheet.Cells[i + skuStart.Row, qtyStart.Column]];
			outputsheet.PageSetup.PrintArea = print_rng.Address;

			return outputsheet;

		}

	}

}
