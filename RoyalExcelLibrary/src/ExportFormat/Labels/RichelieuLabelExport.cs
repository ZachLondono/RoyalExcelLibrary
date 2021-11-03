using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using RoyalExcelLibrary.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.ExportFormat.Labels {
	public class RichelieuLabelExport : ILabelExport {

		private readonly string boxTemplate = "R:\\DB ORDERS\\Labels\\DBox Label Richelieu.label";
		private readonly string shippingTemplate = "R:\\DB ORDERS\\Labels\\Shipping Richelieu Logo.label";

		public void PrintLables(Order order) {


			DymoLabelService boxLabelService = new DymoLabelService(boxTemplate);

			var job = order.Job;
			var boxes = order.Products.Cast<DrawerBox>()
									.OrderByDescending(b => b.Width)
									.OrderByDescending(b => b.Depth);

			int i = 1;
			foreach (var box in boxes) {

				string height = HelperFuncs.FractionalImperialDim(box.Height);
				string width = HelperFuncs.FractionalImperialDim(box.Width);
				string depth = HelperFuncs.FractionalImperialDim(box.Depth);
				string sizeStr = $"{height}\"Hx{width}\"Wx{depth}\"D";

				var label = boxLabelService.CreateLabel();
				label["JOB"] = job.Name;
				label["PO"] = order.Number;
				label["SIZE"] = sizeStr;
				label["QTY"] =	box.Qty;
				label["DESC"] = box.InfoFields[1];
				label["ORDER"] = order.InfoFields[1] + $" : {i}";
				label["NOTE"] = box.InfoFields[0];

				boxLabelService.AddLabel(label, box.Qty);
				i++;
			}

			boxLabelService.PrintLabels();

			DymoLabelService shippingLabelService = new DymoLabelService(shippingTemplate);
			Label shippinglabel = shippingLabelService.CreateLabel();
			shippinglabel["TEXT"] = order.CustomerName;
			shippinglabel["TEXT_1"] = order.InfoFields[0]; // LastName, FirstName
			shippinglabel["TEXT_2"] = order.Number; // Richelieu PO
			shippinglabel["ADDRESS"] = order.ShipAddress.ToString();
			shippingLabelService.AddLabel(shippinglabel, 1);
			shippingLabelService.PrintLabels();


		}

	}

}
