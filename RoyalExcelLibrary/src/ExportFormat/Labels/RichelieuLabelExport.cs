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

		private static readonly string boxTemplate = "R:\\DB ORDERS\\Labels\\DBox Label Richelieu.label";
		private static readonly string shippingTemplate = "R:\\DB ORDERS\\Labels\\Shipping Richelieu Logo.label";

		public void PrintLables(Order order, ILabelServiceFactory factory) {

			RichelieuOrder richOrder = order as RichelieuOrder;

			ILabelService boxLabelService = factory.CreateService(boxTemplate);

			var job = order.Job;
			var boxes = order.Products.Cast<DrawerBox>()
									.OrderByDescending(b => b.Width)
									.OrderByDescending(b => b.Depth);

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
				label["DESC"] = box.ProductDescription;
				label["ORDER"] = richOrder.RichelieuNumber + $" : {box.LineNumber}";
				label["NOTE"] = box.Note;

				boxLabelService.AddLabel(label, box.Qty);
			}

			boxLabelService.PrintLabels();

			ILabelService shippingLabelService = factory.CreateService(shippingTemplate);
			Label shippinglabel = shippingLabelService.CreateLabel();
			shippinglabel["TEXT"] = order.Customer.Name;
			shippinglabel["TEXT_1"] = $"{richOrder.ClientLastName}, {richOrder.ClientFirstName}"; // LastName, FirstName
			shippinglabel["TEXT_2"] = order.Number; // Richelieu PO
			shippinglabel["ADDRESS"] = order.Customer.Address.ToString();
			shippingLabelService.AddLabel(shippinglabel, 1);
			shippingLabelService.PrintLabels();


		}

		public static void PrintSingleRichelieuLabel(int copies, string jobName, string orderNum, string size, string qty, string description, string richOrder, string note, string lineNum) {

			DymoLabelService boxLabelService = new DymoLabelService(boxTemplate);
			var label = boxLabelService.CreateLabel();
			label["JOB"] = jobName;
			label["PO"] = orderNum;
			label["SIZE"] = size;
			label["QTY"] = qty;
			label["DESC"] = description;
			label["ORDER"] = richOrder + " : " + lineNum;
			label["NOTE"] = note;

			boxLabelService.AddLabel(label, copies);
			boxLabelService.PrintLabels();

		}

	}

}
