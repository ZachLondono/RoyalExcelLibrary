using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System;
using System.Linq;

namespace RoyalExcelLibrary.ExportFormat.Google {
	public class HafeleGoogleSheetExport : GoogleSheetsExport {
		public override void ExportOrder(Order order) {
#if DEBUG
			Data.Add("test");
#else
			Data.Add("hafele");
#endif
			Data.Add(order.Job.CreationDate.ToShortDateString());
			Data.Add(order.Number);	// Hafele PO
			Data.Add(order.InfoFields[1]); // Hafele Project
			Data.Add(""); // CFG #
			Data.Add(order.CustomerName); // Customer Name
			Data.Add(order.Job.Name); // Customer PO

			int totalDBCount = order.Products.Where(p => p is DrawerBox)
											.Select(p => (p as DrawerBox).Qty)
											.Sum();

			Data.Add(totalDBCount == 0 ? "" : totalDBCount.ToString());
			Data.Add(DateTime.Now.AddDays(7).ToShortDateString()); // Ship Date
			Data.Add(order.ShippingCost + order.SubTotal);
			Data.Add(order.InfoFields[2]); // Pro Number

			ExportCurrentData();
		}

	}

}
