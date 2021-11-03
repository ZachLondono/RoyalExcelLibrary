using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using RoyalExcelLibrary.Services;
using System.Linq;

namespace RoyalExcelLibrary.ExportFormat.Labels {
	public class HafeleLabelExport : ILabelExport {

		private readonly string boxTemplate = "R:\\DB ORDERS\\Labels\\HafeleLabel-1.label";
		private readonly string shippingTemplate = "R:\\DB ORDERS\\Labels\\LargeShipping Hafele Logo.label";

		public void PrintLables(Order order) {

			DymoLabelService boxLabelService = new DymoLabelService(boxTemplate);

			var job = order.Job;

			var boxes = order.Products.Cast<DrawerBox>()
									.OrderByDescending(b => b.Width)
									.OrderByDescending(b => b.Depth);

			string cfgNum = order.InfoFields[0];
			string projectNum = order.InfoFields[1];

			int i = 1;
			foreach (var box in boxes) {

				string height = HelperFuncs.FractionalImperialDim(box.Height);
				string width = HelperFuncs.FractionalImperialDim(box.Width);
				string depth = HelperFuncs.FractionalImperialDim(box.Depth);
				string sizeStr = $"{height}\"Hx{width}\"Wx{depth}\"D";

				string jobName = box.InfoFields[0];
				string note = box.InfoFields[1];

				var label = boxLabelService.CreateLabel();
				label["CustomerName"] = order.CustomerName;
				label["ClientPO"] = job.Name;
				label["HafelePO"] = order.Number;
				label["CFG"] = cfgNum;
				label["JobName"] = jobName;
				label["Qty"] = $"{box.Qty}";
				label["LineNum"] = $"{i}";
				label["Size"] = sizeStr;
				label["Message"] = note;

				boxLabelService.AddLabel(label, box.Qty);

			}

			boxLabelService.PrintLabels();

			DymoLabelService shippingLabelService = new DymoLabelService(shippingTemplate);
			Label shippinglabel = shippingLabelService.CreateLabel();
			shippinglabel["Company"] = order.CustomerName;
			shippinglabel["PO"] = job.Name;
			shippinglabel["Cfg"] = "";
			shippinglabel["HafelePO"] = order.Number;
			shippinglabel["HafeleProject"] = projectNum;
			shippingLabelService.AddLabel(shippinglabel, 1);
			shippingLabelService.PrintLabels();


		}

	}

}
