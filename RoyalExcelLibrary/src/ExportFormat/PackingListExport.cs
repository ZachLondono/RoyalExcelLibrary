using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace RoyalExcelLibrary.ExportFormat {
	public class PackingListExport : IExcelExport {

		public readonly string _packingListTemplateFile = "R:\\DB ORDERS\\RoyalExcelLibrary\\Export Templates\\PackingListTemplate.xlsx";

		public Worksheet ExportOrder(Order order, Workbook workbook) {

			Worksheet outputsheet;
			string worksheetname = "Packing List";

			outputsheet = HelperFuncs.LoadTemplate(_packingListTemplateFile, worksheetname, workbook);

			Company supplierDetails = order.Supplier;

			Range supplier = outputsheet.Range["SupplierName"];
			supplier.Value2 = supplierDetails.Name;
			Range supplierAddress = outputsheet.Range["SupplierAddress"];
			supplierAddress.Value2 = supplierDetails.Address.Line1;
			Range supplierAddress2 = outputsheet.Range["SupplierAddress2"];
			supplierAddress2.Value2 = $"{supplierDetails.Address.City}, {supplierDetails.Address.State}, {supplierDetails.Address.Zip}";

			Company customerDetails = order.Customer;

			Range recipient = outputsheet.Range["RecipientName"];
			recipient.Value2 = customerDetails.Name;
			Range recipientAddress = outputsheet.Range["RecipientAddress"];
			recipientAddress.Value2 = customerDetails.Address?.Line1 ?? "";
			Range recipientAddress2 = outputsheet.Range["RecipientAddress2"];
			recipientAddress2.Value2 = $"{customerDetails.Address?.City ?? ""}, {customerDetails.Address?.State ?? ""}, {customerDetails.Address?.Zip ?? ""}";

			Range date = outputsheet.Range["Date"];
			date.Value2 = DateTime.Today.ToShortDateString();

			Range label1 = outputsheet.Range["Label1"];
			Range label2 = outputsheet.Range["Label2"];
			Range label3 = outputsheet.Range["Label3"];
			Range label4 = outputsheet.Range["Label4"];

			Range value1 = outputsheet.Range["Value1"];
			Range value2 = outputsheet.Range["Value2"];
			Range value3 = outputsheet.Range["Value3"];
			Range value4 = outputsheet.Range["Value4"];

			if (order.Job.JobSource.ToLower().Equals("allmoxy")) {
				label2.Value2 = "Order Name";
				value2.Value2 = order.Job.Name;

				label1.Value2 = "Allmoxy #";
				value1.Value2 = order.Number;
			}

			Range lineNumStart = outputsheet.Range["LineNumStart"];
			Range qtyStart = outputsheet.Range["QtyStart"];
			Range descStart = outputsheet.Range["DescriptionStart"];
			Range logoStart = outputsheet.Range["LogoStart"];
			Range heightStart = outputsheet.Range["HeightStart"];
			Range widthStart = outputsheet.Range["WidthStart"];
			Range depthStart = outputsheet.Range["DepthStart"];

			IEnumerable<DrawerBox> boxes = order.Products.Where(p => p is DrawerBox).Cast<DrawerBox>();

			int i = 0;
			foreach (DrawerBox box in boxes) {
				lineNumStart.Offset[i, 0].Value2 = i + 1;
				qtyStart.Offset[i, 0].Value2 = box.Qty;
				descStart.Offset[i, 0].Value2 = box is UDrawerBox ? "U-Shaped Dovetail Drawer Box" : "Dovetail Drawer Box";
				logoStart.Offset[i, 0].Value2 = box.Logo ? "Yes" : "";
				heightStart.Offset[i, 0].Value2 = HelperFuncs.FractionalImperialDim(box.Height);
				widthStart.Offset[i, 0].Value2 = HelperFuncs.FractionalImperialDim(box.Width);
				depthStart.Offset[i, 0].Value2 = HelperFuncs.FractionalImperialDim(box.Depth);
				i++;
			}

			int boxCount = boxes.Sum(b => b.Qty);

			outputsheet.Range["ItemCount"].Value2 = boxCount;

			return outputsheet;

		}
	}

}
