using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System;
using System.Linq;

namespace RoyalExcelLibrary.ExportFormat.Google {
	public class RichelieuGoogleSheetExport : GoogleSheetsExport {

		public override void ExportOrder(Order order) {

			RichelieuOrder richOrder = order as RichelieuOrder;

#if DEBUG
			Data.Add("test");
#else
			Data.Add("richelieu");
#endif
			Data.Add(richOrder.WebNumber); // Web Number
			Data.Add(order.Number); // Rich Order Number
			Data.Add(richOrder.RichelieuNumber); // Rich PO
			Data.Add(DateTime.Now.AddDays(7).ToShortDateString()); // Pickup Date
			Data.Add(order.Customer.Name); // Customer

			int totalDBCount = order.Products.Where(p => p is DrawerBox)
											.Select(p => (p as DrawerBox).Qty)
											.Sum();

			Data.Add(totalDBCount); // Box count
			Data.Add(order.SubTotal + order.ShippingCost); // Invoice amount

			ExportCurrentData();
		}

	}

}