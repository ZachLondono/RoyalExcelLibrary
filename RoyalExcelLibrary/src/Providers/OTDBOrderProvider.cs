using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System.Diagnostics;

namespace RoyalExcelLibrary.Providers {
	public class OTDBOrderProvider : IOrderProvider {
	
		private Excel.Application _app { get; set; }
		private UnitType _units { get; set; }

		public OTDBOrderProvider(Excel.Application app) {
			_app = app;
			_units = UnitType.Inches;
		}

		public OTDBOrderProvider(Excel.Application app, UnitType units) {
			_app = app;
			_units = units;
		}

		// <summary>
		// Loads the current job from an excel workbook which follows the OnTrack excel workbook format
		// </summary>
		// <excepiton cref="ArgumentOutOfRangeException">Thrown when unable to find required ranges in the excel sheet</exception>
		public Order LoadCurrentOrder() {

			string Id = TryGetRange("OrderName").Value2.ToString();
			string MatStr = TryGetRange("Material").Value2.ToString();
			string BotMatStr = TryGetRange("BotThickness").Value2.ToString();

			double grossRevenue = TryGetRange("R4").Value2;

			Job job = new Job();
			job.Name = Id;
			job.CreationDate = DateTime.Now;
			job.GrossRevenue = grossRevenue;

			MaterialType sideMat = ParseMaterial(MatStr);
			MaterialType bottomMat = ParseMaterial(BotMatStr);

			Excel.Range qtyStart = _app.Range["B16"];
			Excel.Range heightStart = _app.Range["C16"];
			Excel.Range widthStart = _app.Range["D16"];
			Excel.Range depthStart = _app.Range["E16"];

			bool convertToMM = _units == UnitType.Millimeters ? false : true;

			List<DrawerBox> boxes = new List<DrawerBox>();

			int maxCount = 200;
			int i = 0;
			while (i < maxCount) {

				try {

					Excel.Range qty = qtyStart.Offset[i,0];
					if (qty.Value2 is null || string.IsNullOrEmpty(qty.Value2.ToString()))
						break;

					DrawerBox box = new DrawerBox();
					box.SideMaterial = sideMat;
					box.BottomMaterial = bottomMat;

					box.Qty = Convert.ToInt32(qty.Value2);
					box.Height = Convert.ToDouble(heightStart.Offset[i, 0].Value2) * (convertToMM ? 25.4 : 1);
					box.Width = Convert.ToDouble(widthStart.Offset[i, 0].Value2) * (convertToMM ? 25.4 : 1);
					box.Depth = Convert.ToDouble(depthStart.Offset[i, 0].Value2) * (convertToMM ? 25.4 : 1);

					Debug.WriteLine($"q{box.Qty}: {box.Height}x{box.Width}x{box.Depth}");

					boxes.Add(box);

				} catch (Exception e) {
					Debug.WriteLine($"Unable to parse box on line #{i}\n{e}");
				}

				i++;
			}

			Order order = new Order(job);
			order.AddProducts(boxes);

			return order;

		}

		private Excel.Range TryGetRange(string name) {
			Excel.Range range = _app.Range[name];
			if (range is null)
				throw new ArgumentOutOfRangeException("name", name, $"Unable to access range '{name}'");
			return range;
		}

		private MaterialType ParseMaterial(string name) {

			switch (name) {
				case "Economy Birch":
					return MaterialType.EconomyBirch;
				case "Solid Birch":
					return MaterialType.SolidBirch;
				case "Hybrid":
					return MaterialType.HybridBirch;
				case "Walnut":
					return MaterialType.SolidWalnut;
				case "1/4\" Plywood":
					return MaterialType.Plywood1_4;
				case "1/2\" Plywood":
					return MaterialType.Plywood1_2;
				default:
					return MaterialType.Unknown;
			}

		}

	}
}
