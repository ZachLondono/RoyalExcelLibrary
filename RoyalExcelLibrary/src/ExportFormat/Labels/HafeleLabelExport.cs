﻿using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using RoyalExcelLibrary.Services;
using System.Linq;

namespace RoyalExcelLibrary.ExportFormat.Labels {
	public class HafeleLabelExport : ILabelExport {

		private readonly string boxTemplate = "R:\\DB ORDERS\\Labels\\HafeleLabel-1.label";
		private readonly string shippingTemplate = "R:\\DB ORDERS\\Labels\\LargeShipping Hafele Logo.label";
		public string ProjectNum { get; set; }

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
				label["CustomerName"] = order.CustomerName;
				label["ClientPO"] = job.Name;
				label["HafelePO"] = order.Number;
				label["CFG"] = box.LabelNote;
				label["JobName"] = "";
				label["Qty"] = $"{box.Qty}";
				label["LineNum"] = $"{i}";
				label["Size"] = sizeStr;
				label["Message"] = "Made in USA";

				boxLabelService.AddLabel(label, box.Qty);

			}

			boxLabelService.PrintLabels();

			DymoLabelService shippingLabelService = new DymoLabelService(shippingTemplate);
			Label shippinglabel = shippingLabelService.CreateLabel();
			shippinglabel["Company"] = order.CustomerName;
			shippinglabel["PO"] = job.Name;
			shippinglabel["Cfg"] = "";
			shippinglabel["HafelePO"] = order.Number;
			shippinglabel["HafeleProject"] = ProjectNum;
			shippingLabelService.AddLabel(shippinglabel, 1);
			shippingLabelService.PrintLabels();


		}

	}

}