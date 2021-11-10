using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace RoyalExcelLibrary.ExportFormat {
	public class InvoiceExport : IExcelExport {

		public readonly string _invoiceTemplate = "R:\\DB ORDERS\\RoyalExcelLibrary\\Export Templates\\InvoiceTemplate.xlsx";

		public Worksheet ExportOrder(Order order, Workbook workbook) {

			Worksheet outputsheet;
			string worksheetname = "Invoice";

			try {
				outputsheet = workbook.Worksheets[worksheetname];
			} catch (COMException) {
				// TODO copy packing list from template workbook
				Application app = (Application)ExcelDnaUtil.Application;
				Workbook template = app.Workbooks.Open(_invoiceTemplate);
				template.Worksheets[worksheetname].Copy(workbook.Worksheets[workbook.Worksheets.Count - 1]);
				template.Close(SaveChanges: false);
				outputsheet = workbook.Worksheets[worksheetname];
			}

			Range supplier = outputsheet.Range["SupplierName"];
			supplier.Value2 = order.Supplier.Name;
			Range supplierAddress = outputsheet.Range["SupplierAddress"];
			supplierAddress.Value2 = order.Supplier.Address.Line1;
			Range supplierAddress2 = outputsheet.Range["SupplierAddress2"];
			supplierAddress2.Value2 = $"{order.Supplier.Address.City}, {order.Supplier.Address.State}, {order.Supplier.Address.Zip}";

			Company invoiceRecipient = order is HafeleOrder ? order.Vendor : order.Customer;

			Range recipient = outputsheet.Range["RecipientName"];
			recipient.Value2 = invoiceRecipient.Name;
			Range recipientAddress = outputsheet.Range["RecipientAddress"];
			recipientAddress.Value2 = invoiceRecipient.Address?.Line1 ?? "";
			Range recipientAddressLine2 = outputsheet.Range["RecipientAddressLine2"];
			recipientAddressLine2.Value2 = invoiceRecipient.Address.Line2;
			Range recipientAddress2 = outputsheet.Range["RecipientAddress2"];
			recipientAddress2.Value2 = $"{invoiceRecipient.Address?.City ?? ""}, {invoiceRecipient.Address?.State ?? ""}, {invoiceRecipient.Address?.Zip ?? ""}";

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
			} else if (order.Job.JobSource.ToLower().Equals("hafele")) {

				HafeleOrder hafeleOrder = order as HafeleOrder;

				label1.Value2 = "Shipping Number";
				value1.Value2 = hafeleOrder.ProNumber;
				label2.Value2 = "Hafele Project";
				value2.Value2 = hafeleOrder.ProjectNumber;
				label3.Value2 = "Customer PO";
				value3.Value2 = hafeleOrder.ClientPurchaseOrder;
				label4.Value2 = "Customer Name";
				value4.Value2 = hafeleOrder.Customer.Name;

            }

			Range refNum = outputsheet.Range["RefNum"];
			refNum.Value2 = order.Number;

			Range lineNumStart = outputsheet.Range["LineNumStart"];
			Range qtyStart = outputsheet.Range["QtyStart"];
			Range descStart = outputsheet.Range["DescriptionStart"];
			Range logoStart = outputsheet.Range["LogoStart"];
			Range heightStart = outputsheet.Range["HeightStart"];
			Range widthStart = outputsheet.Range["WidthStart"];
			Range depthStart = outputsheet.Range["DepthStart"];
			Range priceStart = outputsheet.Range["PriceStart"];
			Range extPriceStart = outputsheet.Range["ExtPriceStart"];

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
				priceStart.Offset[i, 0].Value2 = box.UnitPrice;
				extPriceStart.Offset[i, 0].Value2 = box.UnitPrice * box.Qty;
				i++;
			}

			int lastRow = qtyStart.Row + i;
			int lastCol = extPriceStart.Column;

			decimal invoiceTotal = boxes.Sum(b => b.UnitPrice * b.Qty);
			outputsheet.Range["InvoiceTotal"].Value2 = invoiceTotal;

			int boxCount = boxes.Sum(b => b.Qty);
			outputsheet.Range["ItemCount"].Value2 = boxCount;

			Range print_rng = outputsheet.Range[outputsheet.Cells[1, 1], outputsheet.Cells[lastRow, lastCol]];
			outputsheet.PageSetup.PrintArea = print_rng.Address;

			return outputsheet;

		}
	}


}