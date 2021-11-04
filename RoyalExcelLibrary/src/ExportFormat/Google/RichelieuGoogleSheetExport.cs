using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System;
using System.Linq;

namespace RoyalExcelLibrary.ExportFormat.Google {
	public class RichelieuGoogleSheetExport : GoogleSheetsExport {

		public override void ExportOrder(Order order) {
#if DEBUG
			Data.Add("test");
#else
			Data.Add("richelieu");
#endif
			Data.Add(order.InfoFields[2]); // Web Number
			Data.Add(order.InfoFields[1]); // Rich Order Number
			Data.Add(order.Number); // Rich PO
			Data.Add(DateTime.Now.AddDays(7).ToShortDateString()); // Pickup Date
			Data.Add(order.CustomerName); // Customer

			int totalDBCount = order.Products.Where(p => p is DrawerBox)
											.Select(p => (p as DrawerBox).Qty)
											.Sum();

			Data.Add(totalDBCount); // Box count
			Data.Add(order.SubTotal + order.ShippingCost); // Invoice amount

			ExportCurrentData();
		}

	}

}