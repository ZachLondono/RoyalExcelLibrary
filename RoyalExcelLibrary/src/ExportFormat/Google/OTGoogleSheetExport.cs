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


			// For standard OT orders that come from harold, OT processes the payments and therefore owes royal the cost of the boxes
			// For OT orders that come through Allmoxy

			decimal total = order.ShippingCost + order.SubTotal;
			decimal totalPaid = 0;
			decimal commissionRate = 0.13M;
			decimal stripeFee = 0;
			if (order.Job.JobSource.ToLower().Equals("allmoxy")) {
				commissionRate = 0.1M;
				totalPaid = total;
				stripeFee = ExcelLibrary.CalculateStripeFee(total);
				total += order.Tax;
			}

			decimal commission = ExcelLibrary.CalculateCommissionPayment(total, order.ShippingCost, order.Tax, stripeFee, commissionRate);

			Data.Add(total);
			Data.Add(total - totalPaid - commission - order.ShippingCost);

			ExportCurrentData();
		}

	}

	public class MetroGoogleSheetExport : GoogleSheetsExport {

		public override void ExportOrder(Order order) {
#if DEBUG
			Data.Add("test");
#else
			Data.Add("metro");
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
			Data.Add(order.Tax);
			decimal total = order.ShippingCost + order.SubTotal + order.Tax;
			Data.Add(total);
			decimal stripeFee = ExcelLibrary.CalculateStripeFee(total);
			Data.Add(stripeFee);

			ExportCurrentData();
		}

	}

}
