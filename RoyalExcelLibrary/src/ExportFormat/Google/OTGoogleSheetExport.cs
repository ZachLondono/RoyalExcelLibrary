using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System;
using System.Linq;

namespace RoyalExcelLibrary.ExportFormat.Google {
	public class OTGoogleSheetExport : GoogleSheetsExport {

		public override void ExportOrder(Order order) {
#if DEBUG
			Data.Add("test");
#else
			Data.Add("ot");
#endif
			Data.Add(order.Number);
			Data.Add(order.Job.CreationDate.ToShortDateString());
			Data.Add(order.CustomerName);
			Data.Add(order.Job.Name);

			int totalDBCount = order.Products.Where(p => p is DrawerBox)
											.Select(p => (p as DrawerBox).Qty)
											.Sum();

			int totalDoorCount = 0;

			Data.Add(totalDBCount == 0 ? "" : totalDBCount.ToString());
			Data.Add(totalDoorCount == 0 ? "" : totalDoorCount.ToString());
			Data.Add("");
			Data.Add("");
			Data.Add(order.SubTotal);
			Data.Add(order.ShippingCost);
			Data.Add(order.ShippingCost + order.SubTotal);
			Data.Add(order.SubTotal * 0.87);

			ExportCurrentData();
		}

	}

}
