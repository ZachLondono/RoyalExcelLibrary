using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

using RoyalExcelLibrary.ExcelUI.Models;
using RoyalExcelLibrary.ExcelUI.Models.Products;
using System.Diagnostics;
using RoyalExcelLibrary.ExcelUI.Models.Options;

namespace RoyalExcelLibrary.ExcelUI.Providers {
	public class OTDBOrderProvider : IExcelOrderProvider {

		public Excel.Application App { get;  set; }
		private readonly UnitType _units;
		private readonly AppSettings _settings;

		public OTDBOrderProvider() {
			_units = UnitType.Inches;
			_settings = HelperFuncs.ReadSettings();
		}

		public OTDBOrderProvider(Excel.Application app, UnitType units) {
			App = app;
			_units = units;
			_settings = HelperFuncs.ReadSettings();
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
			string clips = TryGetRange("C7").Value2?.ToString() ?? "";
			string postFinishStr = TryGetRange("C8").Value2?.ToString() ?? "";
			decimal grossRevenue = Convert.ToDecimal(TryGetRange("R4").Value2 ?? "0");
            
            DateTime orderDate;
            try {
                orderDate = DateTime.Parse(TryGetRange("Date").Value2.ToString());
            } catch {
                orderDate = DateTime.Today;
            }

            Job job = new Job {
                JobSource = "OT",
                Name = jobName,
                CreationDate = orderDate,
                GrossRevenue = grossRevenue
            };

            string sideMat = ParseMaterial(MatStr);
			string bottomMat = ParseMaterial(BotMatStr);
			UndermountNotch notch = ParseNotch(notchStr);
			bool postFinish = postFinishStr.Equals("Yes");

			Excel.Range qtyStart = App.Range["B16"];
			Excel.Range heightStart = App.Range["C16"];
			Excel.Range widthStart = App.Range["D16"];
			Excel.Range depthStart = App.Range["E16"];
			Excel.Range noteStart = App.Range["R16"];
			Excel.Range pulloutStart = App.Range["F16"];
			Excel.Range logoStart = App.Range["J16"];
			Excel.Range accessoryStart = App.Range["K16"];
			Excel.Range aStart = App.Range["T16"];
			Excel.Range bStart = App.Range["U16"];
			Excel.Range cStart = App.Range["V16"];


			bool convertToMM = _units != UnitType.Millimeters;

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
					if (accessoryStr.Equals("U-Box")) {
						box = new UDrawerBox();
						(box as UDrawerBox).A = Convert.ToDouble(aStart.Offset[i, 0].Value2) * (convertToMM ? 25.4 : 1);
						(box as UDrawerBox).B = Convert.ToDouble(bStart.Offset[i, 0].Value2) * (convertToMM ? 25.4 : 1);
						(box as UDrawerBox).C = Convert.ToDouble(cStart.Offset[i, 0].Value2) * (convertToMM ? 25.4 : 1);
					}  else box = new DrawerBox();

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
					box.InsertOption = accessoryStr;
					box.Logo = logoStart.Offset[i, 0].Value2?.Equals("Yes") ?? false;
					box.ScoopFront = pulloutStart.Offset[i,0].Value2?.Equals("Scoop Front") ?? false;
					box.LineNumber = lineNum++;
					box.Note = noteStart.Offset[i, 0].Value2?.ToString() ?? "";

					Debug.WriteLine($"q{box.Qty}: {box.Height}x{box.Width}x{box.Depth}");

					boxes.Add(box);

				} catch (Exception e) {
					Debug.WriteLine($"Unable to parse box on line #{i}\n{e}");
				}
				i++;
			}
			
			string customer = TryGetRange("CustomerName").Value2.ToString();
			string orderNum = TryGetRange("OrderName").Value2.ToString();
			string vendorName = TryGetRange("VendorName").Value2.ToString();

			string addressLine1;
			string addressLine2;
			string city;
			string state;
			string zip;
			TryReadRangeValue("Address1", out addressLine1);
			TryReadRangeValue("Address2", out addressLine2);
			TryReadRangeValue("City", out city);
			TryReadRangeValue("State", out state);
			TryReadRangeValue("Zip", out zip);

			string orderNotes = "";

			try {
				var noteRng = App.Range["OrderNotes"];
				if (!(noteRng is null)) {
					orderNotes = noteRng.Value2.ToString();
				}
			} catch { }

			try {
				var sideOptionRng = App.Range["SideOption"];
				if (!(sideOptionRng is null)) {
					var sideOption = sideOptionRng.Value2;
					if (!string.IsNullOrEmpty(orderNotes))
						orderNotes += " | ";
					orderNotes = $"Sides: {sideOption}";
				}
			} catch { }


			Order order = new Order(job);
			order.AddProducts(boxes);
			order.Number = orderNum;
			order.ShippingCost = Convert.ToDecimal(TryGetRange("R7").Value2 ?? "0");
			order.SubTotal = Convert.ToDecimal(TryGetRange("Invoice!I8").Value2 ?? "0") - order.ShippingCost;
			order.Comment = orderNotes;
			order.Customer = new Company {
				Name = customer,
				Address = new ExportFormat.Address {
					Line1 = addressLine1,
					Line2 = addressLine2,
					City = city,
					State = state,
					Zip = zip
                }
			};
			order.Vendor = new Company {
				Name = vendorName,
				Address = new ExportFormat.Address {
					Line1 = "",
					Line2 = "",
					City = "",
					State = "",
					Zip = ""
				}
			};

			return order;

		}

		private bool TryReadRangeValue(string rangeName, out string value) {
			try {
				Excel.Range range = App.Range[rangeName];
				if (range is null) {
					value = "";
					return false;
				}

				value = range.Value2?.ToString() ?? "";
				return true;
			} catch {
				value = "";
				return false;
			}
		}

		private Excel.Range TryGetRange(string name) {
			Excel.Range range = App.Range[name];
			if (range is null)
				throw new ArgumentOutOfRangeException("name", name, $"Unable to access range '{name}'");
			return range;
		}

		private string ParseMaterial(string name) {
			var matMap = _settings.MaterialProfiles["ot"];
			if (matMap is null) return name;
			return matMap[name];
		}

		private UndermountNotch ParseNotch(string name) {

			switch (name) {
				case "Notch for U/M Slide":
					return UndermountNotch.Std_Notch;
				case "Notch for U/M Slide-Wide":
					return UndermountNotch.Wide_Notch;
				case "Notch Front & Back":
					return UndermountNotch.Front_Back;
				case "Notch for 828 Slides":
					return UndermountNotch.Notch_828;
				case "":
				case "No Notch":
					return UndermountNotch.No_Notch;
				default:
					return UndermountNotch.Unknown;
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
