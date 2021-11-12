using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using RoyalExcelLibrary.Services;
using System.Linq;

namespace RoyalExcelLibrary.ExportFormat.Labels {
	public class OTLabelExport : ILabelExport {

		private static readonly string boxTemplate = "R:\\DB ORDERS\\Labels\\DBox Label - OT Large.label";

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

				string note = box.Note;

				var label = boxLabelService.CreateLabel();
				label["Name"] = order.Customer.Name;
				label["Size"] = sizeStr;
				label["QTY"] = $"{box.Qty}";
				label["ID"] = $"{order.Number} - {box.LineNumber}";
				label["Job"] = note;
				label["Number"] = job.Name;

				boxLabelService.AddLabel(label, box.Qty);

			}

			boxLabelService.PrintLabels();

		}

		public static void PrintSingleOTLabel(int copies, string customerName, string size, string qty, string orderNumber, string lineNum, string note, string jobName) {
			
			DymoLabelService boxLabelService = new DymoLabelService(boxTemplate);

			var label = boxLabelService.CreateLabel();
			label["Name"] = customerName;
			label["Size"] = size;
			label["QTY"] = qty;
			label["ID"] = orderNumber + " - " + lineNum;
			label["Job"] = note;
			label["Number"] = jobName;

			boxLabelService.AddLabel(label, copies);

			boxLabelService.PrintLabels();

		}

	}

}
