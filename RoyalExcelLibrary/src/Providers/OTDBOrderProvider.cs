using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System.Diagnostics;
using RoyalExcelLibrary.Models.Options;

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

			string jobName = TryGetRange("JobName").Value2.ToString();
			string MatStr = TryGetRange("Material").Value2.ToString();
			string BotMatStr = TryGetRange("BotThickness").Value2.ToString();
			string notchStr = TryGetRange("Notch").Value2.ToString();
			string clipsStr = TryGetRange("C7").Value2?.ToString() ?? "";
			string postFinishStr = TryGetRange("C8").Value2?.ToString() ?? "";
			double grossRevenue = TryGetRange("R4").Value2 ?? "";

			Job job = new Job();
			job.JobSource = "OT";
			job.Status = Status.UnConfirmed;
			job.Name = jobName;
			job.CreationDate = DateTime.Now;
			job.GrossRevenue = grossRevenue;

			MaterialType sideMat = ParseMaterial(MatStr);
			MaterialType bottomMat = ParseMaterial(BotMatStr);
			UndermountNotch notch = ParseNotch(notchStr);
			Clips clips = ParseClips(clipsStr);
			bool postFinish = postFinishStr.Equals("Yes");

			Excel.Range qtyStart = _app.Range["B16"];
			Excel.Range heightStart = _app.Range["C16"];
			Excel.Range widthStart = _app.Range["D16"];
			Excel.Range depthStart = _app.Range["E16"];
			Excel.Range noteStart = _app.Range["R16"];
			Excel.Range pulloutStart = _app.Range["F16"];
			Excel.Range logoStart = _app.Range["J16"];
			Excel.Range accessoryStart = _app.Range["K16"];

			bool convertToMM = _units == UnitType.Millimeters ? false : true;

			List<DrawerBox> boxes = new List<DrawerBox>();

			int lineNum = 1;
			int maxCount = 200;
			int i = 0;
			while (i < maxCount) {

				try {

					Excel.Range qty = qtyStart.Offset[i,0];
					if (qty.Value2 is null || string.IsNullOrEmpty(qty.Value2.ToString()))
						break;

					string accessoryStr = accessoryStart.Offset[i, 0].Value2?.ToString() ?? "";

					DrawerBox box;
					if (accessoryStr.Equals("U-Box"))
						box = new UDrawerBox();
					else box = new DrawerBox();

					box.SideMaterial = sideMat;
					box.BottomMaterial = bottomMat;
					box.NotchOption = notch;
					box.ClipsOption = clips;
					box.PostFinish = postFinish;
					box.MountingHoles = false;

					box.Qty = Convert.ToInt32(qty.Value2);
					box.Height = Convert.ToDouble(heightStart.Offset[i, 0].Value2) * (convertToMM ? 25.4 : 1);
					box.Width = Convert.ToDouble(widthStart.Offset[i, 0].Value2) * (convertToMM ? 25.4 : 1);
					box.Depth = Convert.ToDouble(depthStart.Offset[i, 0].Value2) * (convertToMM ? 25.4 : 1);
					box.InsertOption = ParseInsert(accessoryStr);
					box.Logo = logoStart.Offset[i, 0].Value2?.Equals("Yes") ?? false;
					box.ScoopFront = pulloutStart.Offset[i,0].Value2?.Equals("Scoop Front") ?? false;
					box.LineNumber = lineNum++;

					List<string> info = new List<string>();
					if (noteStart.Offset[i, 0].Value2 is null)
						info.Add("");
					else info.Add(noteStart.Offset[i, 0].Value2.ToString());
					box.InfoFields = info;

					Debug.WriteLine($"q{box.Qty}: {box.Height}x{box.Width}x{box.Depth}");

					boxes.Add(box);

				} catch (Exception e) {
					Debug.WriteLine($"Unable to parse box on line #{i}\n{e}");
				}
				i++;
			}
			
			string customer = TryGetRange("CustomerName").Value2.ToString();
			string orderNum = TryGetRange("OrderName").Value2.ToString();

			Order order = new Order(job, customer, orderNum);
			order.AddProducts(boxes);
			order.ShippingCost = Convert.ToDouble(TryGetRange("R7").Value2);
			order.SubTotal = Convert.ToDouble(TryGetRange("Invoice!I8").Value2) - order.ShippingCost;
			

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
				case "White Oak":
					return MaterialType.WhiteOak;
				case "1/4\" Plywood":
					return MaterialType.Plywood1_4;
				case "1/2\" Plywood":
					return MaterialType.Plywood1_2;
				default:
					return MaterialType.Unknown;
			}

		}

		private UndermountNotch ParseNotch(string name) {

			switch (name) {
				case "Notch for U/M Slide":
					return UndermountNotch.Std_Notch;
				case "Notch for U/M Slide-Wide":
					return UndermountNotch.Wide_Notch;
				case "Notch Front & Back":
					return UndermountNotch.Front_Back;
				case "":
				case "No Notch":
					return UndermountNotch.No_Notch;
				default:
					return UndermountNotch.Unknown;
			}

		}

		private Clips ParseClips(string name) {
			switch (name) {
				case "Hafele":
					return Clips.Hafele;
				case "Richelieu":
					return Clips.Richelieu;
				case "Blum":
					return Clips.Blum;
				case "Hettich":
					return Clips.Hettich;
				case "":
				case "None":
					return Clips.No_Clips;
				default:
					return Clips.Unknown;
			}
		}

		private Insert ParseInsert(string name) {
			
			switch (name) {
				case "Cutlery Insert 15\"":
					return Insert.Cutlery_15;
				case "Cutlery Insert 23.5\"":
					return Insert.Cutlery_23;
				case "Fixed Divider 2":
					return Insert.Divider_2;
				case "Fixed Divider 3":
					return Insert.Divider_3;
				case "Fixed Divider 4":
					return Insert.Divider_4;
				case "Fixed Divider 6":
					return Insert.Divider_6;
				case "Fixed Divider 7":
					return Insert.Divider_7;
				case "Fixed Divider 5":
					return Insert.Divider_5;
				case "Fixed Divider 8":
					return Insert.Divider_8;
				case "":
				case "None":
				case "U-Box":
					return Insert.No_Insert;
				default:
					return Insert.Unknown;
			}

		}

	}
}
