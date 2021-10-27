using Microsoft.VisualStudio.TestTools.UnitTesting;
using RoyalExcelLibrary.ExportFormat.Labels;
using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using RoyalExcelLibrary.Services;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelLibTests {
	
	[TestClass]
	public class LabelTest {

		[TestMethod]
		public void TestLabel() {

			string templatePath = "C:\\Users\\Zachary Londono\\Desktop\\TestLabel.dymo";

			DymoLabelService labelService = new DymoLabelService(templatePath);

			var label0 = labelService.CreateLabel();
			label0["ITextObject0"] = "Hello0";
			label0["ITextObject1"] = "World0";

			var label1 = labelService.CreateLabel();
			label1["ITextObject0"] = "Howdy1";
			label1["ITextObject1"] = "Earth1";

			labelService.AddLabel(label0, 1);
			labelService.AddLabel(label1, 1);

			try {
				labelService.PrintLabels();
			} catch (Exception e) {
				Debug.WriteLine("No label printer available");
			}

		}

		[TestMethod]
		public void TestPrintOrder() {

			OTLabelExport labelExport = new OTLabelExport();

			Order order = new Order(new Job() {
				Name = "Job_Name",
				CreationDate = DateTime.Now,
				GrossRevenue = 0,
				JobSource = "OT",
			},
			customerName:"Customer",
			number:"OT999");;

			order.AddProduct(new DrawerBox {
				Height = 105,
				Width = 300,
				Depth = 300,
				Qty = 123,
				LabelNote = "LabelNote"
			});

			labelExport.PrintLables(order);

		}

	}
}
