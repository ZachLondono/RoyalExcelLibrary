using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using RoyalExcelLibrary.Services;
using System.Linq;

namespace RoyalExcelLibrary.ExportFormat.Labels {
	public class HafeleLabelExport : ILabelExport {

		private static readonly string boxTemplate = "R:\\DB ORDERS\\Labels\\HafeleLabel-1.label";
		private static readonly string shippingTemplate = "R:\\DB ORDERS\\Labels\\LargeShipping Hafele Logo.label";

		public void PrintLables(Order order) {

			HafeleOrder hafeleOrder = order as HafeleOrder;

			DymoLabelService boxLabelService = new DymoLabelService(boxTemplate);

			var job = order.Job;

			var boxes = order.Products.Cast<DrawerBox>()
									.OrderByDescending(b => b.Width)
									.OrderByDescending(b => b.Depth);

			string cfgNum = hafeleOrder.ConfigNumber;
			string projectNum = hafeleOrder.ProjectNumber;

			int i = 1;
			foreach (var box in boxes) {

				string height = HelperFuncs.FractionalImperialDim(box.Height);
				string width = HelperFuncs.FractionalImperialDim(box.Width);
				string depth = HelperFuncs.FractionalImperialDim(box.Depth);
				string sizeStr = $"{height}\"Hx{width}\"Wx{depth}\"D";

				string jobName = box.LevelName;
				string note = box.Note;

				var label = boxLabelService.CreateLabel();
				label["CustomerName"] = order.Customer.Name;
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
			shippinglabel["Company"] = order.Customer.Name;
			shippinglabel["PO"] = job.Name;
			shippinglabel["Cfg"] = "";
			shippinglabel["HafelePO"] = order.Number;
			shippinglabel["HafeleProject"] = projectNum;
			shippingLabelService.AddLabel(shippinglabel, 1);
			shippingLabelService.PrintLabels();


		}

		public static void PrintSingleHafeleLabel	(int copies, string customerName, string clientPO, string hafelePO, string cfgNum, string jobName, string qty, string lineNum, string size, string message) {

			DymoLabelService boxLabelService = new DymoLabelService(boxTemplate);
			var label = boxLabelService.CreateLabel();
			label["CustomerName"] = customerName;
			label["ClientPO"] = clientPO;
			label["HafelePO"] = hafelePO;
			label["CFG"] = cfgNum;
			label["JobName"] = jobName;
			label["Qty"] = qty;
			label["LineNum"] = lineNum;
			label["Size"] = size;
			label["Message"] = message;

			boxLabelService.AddLabel(label, copies);
			boxLabelService.PrintLabels();
		}

	}

}
