using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using RoyalExcelLibrary.Services;
using System.Linq;

namespace RoyalExcelLibrary.ExportFormat.Labels {
	public class OTLabelExport : ILabelExport {

		private readonly string boxTemplate = "R:\\DB ORDERS\\Labels\\DBox Label - OT Large.label";

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

				string note = box.InfoFields[0];

				var label = boxLabelService.CreateLabel();
				label["Name"] = order.CustomerName;
				label["Size"] = sizeStr;
				label["QTY"] = $"{box.Qty}";
				label["ID"] = $"{order.Number} - {box.LineNumber}";
				label["Job"] = note;
				label["Number"] = job.Name;

				boxLabelService.AddLabel(label, box.Qty);

			}

			boxLabelService.PrintLabels();

		}

	}

}
