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
	class RichelieuExcelDBOrderSource : IOrderProvider {

		private Excel.Application _app { get; set; }

		public RichelieuExcelDBOrderSource(Excel.Application app) {
			_app = app;
		}

		public Order LoadCurrentOrder() {

			string jobName = TryGetRange("J23").Value2.ToString();
			double grossRevenue = TryGetRange("'Price Calculator'!R11").Value2;

			Job job = new Job {
				Name = jobName,
				GrossRevenue = grossRevenue,
				CreationDate = DateTime.Now
			};

			Excel.Range qtyStart = TryGetRange("'Price Calculator'!O3");
			Excel.Range heightStart = TryGetRange("'Price Calculator'!B3");
			Excel.Range widthStart = TryGetRange("'Price Calculator'!C3");
			Excel.Range depthStart = TryGetRange("'Price Calculator'!D3");
			Excel.Range sideStart = TryGetRange("'Price Calculator'!E3");
			Excel.Range bottomStart = TryGetRange("'Price Calculator'!G3");

			List<DrawerBox> boxes = new List<DrawerBox>();
			
			int maxCount = 200;
			int i = 0;
			while (i < maxCount) {

				try {

					Excel.Range qty = qtyStart.Offset[i, 0];
					if (qty.Value2 is null || string.IsNullOrEmpty(qty.Value2.ToString()))
						break;

					DrawerBox box = new DrawerBox();
					box.SideMaterial = ParseMaterial(sideStart.Offset[i, 0].Value2.ToString());
					box.BottomMaterial = ParseMaterial(bottomStart.Offset[i, 0].Value2.ToString());
					box.Qty = Convert.ToInt32(qty.Value2);

					var heightVal = heightStart.Offset[i, 0].Value2;
					if (heightVal.GetType() == typeof(string))
						box.Height = FractionToDouble(heightVal) * 25.4;
					else box.Height = heightVal * 25.4;

					var widthVal = widthStart.Offset[i, 0].Value2;
					if (widthVal.GetType() == typeof(string))
						box.Width = FractionToDouble(widthVal) * 25.4;
					else box.Width = widthVal * 25.4;

					var depthVal = depthStart.Offset[i, 0].Value2;
					if (depthVal.GetType() == typeof(string))
						box.Depth = FractionToDouble(depthVal) * 25.4;
					else box.Depth = depthVal * 25.4;

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

		private double FractionToDouble(string fraction) {

			string[] parts = fraction.Split(' ', '/');

			double val = Convert.ToDouble(parts[0]);
			if (parts.Length == 3) {

				double numerator = Convert.ToDouble(parts[1]);
				double denomenator = Convert.ToDouble(parts[2]);

				val += numerator / denomenator;

			}

			return val;

		}

		private Excel.Range TryGetRange(string name) {
			Excel.Range range = _app.Range[name];
			if (range is null)
				throw new ArgumentOutOfRangeException("name", name, $"Unable to access range '{name}'");
			return range;
		}

		private MaterialType ParseMaterial(string name) {

			switch (name) {
				case "Economy Birch (Finger Jointed)":
					return MaterialType.EconomyBirch;
				case "Solid Birch (No Finger Joint)":
					return MaterialType.SolidBirch;
				case "SFJ Birch":
					return MaterialType.HybridBirch;
				case "Walnut":
					return MaterialType.SolidWalnut;
				case "1/4\" Bottom":
					return MaterialType.Plywood1_4;
				case "1/2\" Bottom":
					return MaterialType.Plywood1_2;
				default:
					return MaterialType.Unknown;
			}

		}

	}

}