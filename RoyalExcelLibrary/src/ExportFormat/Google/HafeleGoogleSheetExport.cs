using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System;
using System.Linq;

namespace RoyalExcelLibrary.ExportFormat.Google {
	public class HafeleGoogleSheetExport : GoogleSheetsExport {
		public override void ExportOrder(Order order) {

			HafeleOrder hafeleOrder = order as HafeleOrder;

#if DEBUG
			Data.Add("test");
#else
			Data.Add("hafele");
#endif
			Data.Add(hafeleOrder.Job.CreationDate.ToShortDateString());
			Data.Add(hafeleOrder.Number);	// Hafele PO
			Data.Add(hafeleOrder.ProjectNumber); // Hafele Project
			Data.Add(hafeleOrder.ConfigNumber); // CFG #
			Data.Add(hafeleOrder.Customer.Name); // Customer Name
			Data.Add(hafeleOrder.Job.Name); // Customer PO

			int totalDBCount = hafeleOrder.Products
											.Where(p => p is DrawerBox)
											.Select(p => (p as DrawerBox).Qty)
											.Sum();

			Data.Add(totalDBCount == 0 ? "" : totalDBCount.ToString());
			Data.Add(DateTime.Now.AddDays(7).ToShortDateString()); // Ship Date
			Data.Add(order.ShippingCost + order.SubTotal);
			Data.Add(hafeleOrder.ProNumber); // Pro Number

			ExportCurrentData();
		}

	}

}
