using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System.Diagnostics;
using RoyalExcelLibrary.ExportFormat;
using RoyalExcelLibrary.Models.Options;
using ExcelDna.Integration;
using Microsoft.VisualBasic;

namespace RoyalExcelLibrary.Providers {
	public class HafeleDBOrderProvider : IOrderProvider {

		private string _sourcePath { get; set; }
		private Excel.Worksheet _source { get; set; }

		public HafeleDBOrderProvider(string sourcePath) {
			_sourcePath = sourcePath;
		}


		public Order LoadCurrentOrder() {

			Excel.Workbook sourceBook = (ExcelDnaUtil.Application as Excel.Application).Workbooks.Open(_sourcePath, ReadOnly: true);
			_source = sourceBook.Worksheets["Order Sheet"];

			string clientPO = TryGetRange("K6").Value2.ToString();
			string company = TryGetRange("Company").Value2.ToString();
			Address address = new Address {
				StreetAddress = TryGetRange("V5").Value2.ToString(),
				City = TryGetRange("V7").Value2.ToString(),
				State = TryGetRange("V8").Value2.ToString(),
				Zip = TryGetRange("V9").Value2.ToString()
			};

			decimal grossRevenue = (TryGetRange("G13").Value2 - 50) / 1.3M;
			string hafelePO = TryGetRange("K10").Value2.ToString() ;
			string hafeleProjectNum = TryGetRange("K11").Value2.ToString();
			string hafeleCfg = "";

			Job job = new Job {
				JobSource = "Hafele",
				Status = Status.Confirmed,
				Name = clientPO,
				GrossRevenue = grossRevenue,
				CreationDate = DateTime.Now
			};

			string sideMaterialStr = TryGetRange("Material").Value2.ToString();
			MaterialType sideMaterial = ParseMaterial(sideMaterialStr);

			bool mountingHoles = TryGetRange("MountingHoles").Value2.Equals("Yes");
			bool postFinish = TryGetRange("PostFinish").Value2.Equals("Yes");

			Excel.Range qtyStart = _source.Range["B16"];
			Excel.Range heightStart = _source.Range["F16"];
			Excel.Range widthStart = _source.Range["G16"];
			Excel.Range depthStart = _source.Range["H16"];
			Excel.Range scoopStart = _source.Range["I16"];
			Excel.Range bottomStart = _source.Range["J16"];
			Excel.Range notchStart = _source.Range["k16"];
			Excel.Range logoStart = _source.Range["L16"];
			Excel.Range clipsStart = _source.Range["M16"];
			Excel.Range accessoryStart = _source.Range["N16"];
			Excel.Range jobNameStart = _source.Range["O16"];		// For labels
			Excel.Range unitPriceStart = _source.Range["P16"];
			Excel.Range noteStart = _source.Range["S16"];
			Excel.Range aDimStart = _source.Range["U16"];
			Excel.Range bDimStart = _source.Range["V16"];
			Excel.Range cDimStart = _source.Range["W16"];

			bool convertToMM = TryGetRange("Notation").Value2.Equals("Matric") ? false : true;

			List<DrawerBox> boxes = new List<DrawerBox>();

			int lineNum = 1;
			int maxCount = 200;
			int i = 0;
			while (i < maxCount) {

				try {

					Excel.Range qty = qtyStart.Offset[i, 0];
					if (qty.Value2 is null || string.IsNullOrEmpty(qty.Value2.ToString()))
						break;

					DrawerBox box;
					if (accessoryStart.Offset[i, 0].Value2.Equals("U-Box")) {
						box = new UDrawerBox();
						(box as UDrawerBox).A = aDimStart.Offset[i,0].Value2 *(convertToMM ? 25.4 : 1);
						(box as UDrawerBox).B = bDimStart.Offset[i, 0].Value2 * (convertToMM ? 25.4 : 1);
						(box as UDrawerBox).C = cDimStart.Offset[i, 0].Value2 * (convertToMM ? 25.4 : 1);
					} else {
						box = new DrawerBox();
					}

					box.SideMaterial = sideMaterial;
					box.BottomMaterial = ParseMaterial(bottomStart.Offset[i, 0].Value2.ToString());
					box.ClipsOption = ParseClips(clipsStart.Offset[i,0].Value2);
					box.InsertOption = ParseInsert(accessoryStart.Offset[i,0].Value2);
					box.NotchOption = ParseNotch(notchStart.Offset[i,0].Value2);
					box.Qty = Convert.ToInt32(qty.Value2);
					box.Height = Convert.ToDouble(heightStart.Offset[i, 0].Value2) * (convertToMM ? 25.4 : 1);
					box.Width = Convert.ToDouble(widthStart.Offset[i, 0].Value2) * (convertToMM ? 25.4 : 1);
					box.Depth = Convert.ToDouble(depthStart.Offset[i, 0].Value2) * (convertToMM ? 25.4 : 1);
					box.Logo = logoStart.Offset[i, 0].Value2.Equals("Yes");
					box.ScoopFront = scoopStart.Offset[i, 0].Value2.Equals("Scoop Front");
					box.MountingHoles = mountingHoles;
					box.PostFinish = postFinish;
					box.UnitPrice = unitPriceStart.Offset[i,0].Value2 / 1.3;
					box.LineNumber = lineNum++;

					string jobName = jobNameStart.Offset[i, 0].Value2?.ToString() ?? "";
					string note = noteStart.Offset[i, 0].Value2?.ToString() ?? "";
					List<string> info = new List<string>();
					info.Add(jobName);
					info.Add(note);
					box.InfoFields = info;

					boxes.Add(box);

				} catch (Exception e) {
					Debug.WriteLine($"Unable to parse box on line #{i}\n{e}");
				}

				i++;
			}


			string pronum = Interaction.InputBox("Enter Pro Number", "Pro Number", "none", 0, 0);

			Order order = new Order(job, company, hafelePO);
			order.AddProducts(boxes);
			order.ShipAddress = address;
			order.ShippingCost = 50;
			order.Status = Status.Confirmed;
			order.InfoFields = new List<string>() { hafeleCfg, hafeleProjectNum, pronum };

			sourceBook.Close(SaveChanges: false);

			return order;

		}

		private Excel.Range TryGetRange(string name) {
			Excel.Range range = _source.Range[name];
			if (range is null)
				throw new ArgumentOutOfRangeException("name", name, $"Unable to access range '{name}'");
			return range;
		}

		private Clips ParseClips(string name) {
			
			switch (name) {

				case "":
				case "None":
					return Clips.No_Clips;
				//case "Grass":
				//	return Clips.Grass;
				case "Hettich":
					return Clips.Hettich;
				//case "Salice":
				//	return Clips.Salice;
				case "Blum":
					return Clips.Blum;
				case "Hafele":
					return Clips.Hafele;
				default:
					return Clips.Unknown;
			}

		}

		private UndermountNotch ParseNotch(string name) {
			switch (name) {
				case "":
				case "No Notch":
					return UndermountNotch.No_Notch;
				case "Notch for U/M Slide":
					return UndermountNotch.Std_Notch;
				case "Notch for U/M Slide Wide":
					return UndermountNotch.Wide_Notch;
				case "Notch 828 Slide Front & Back":
					return UndermountNotch.Notch_828;
				default:
					return UndermountNotch.Unknown;
			}
		}

		private Insert ParseInsert(string name) {
			switch (name) { 
				case "Fixed Divider 2":
					return Insert.Divider_2;
				case "Fixed Divider 3":
					return Insert.Divider_3;
				case "Fixed Divider 4":
					return Insert.Divider_4;
				case "Fixed Divider 5":
					return Insert.Divider_5;
				case "Fixed Divider 6":
					return Insert.Divider_6;
				case "Fixed Divider 7":
					return Insert.Divider_7;
				case "Fixed Divider 8":
					return Insert.Divider_8;
				case "Docking Drawer Cutout":
					return Insert.Docking_Cutout;
				case "4\" Dia. Lock Cutout":
					return Insert.Dia_Lock_Cutout;
				case "4\" Dia. Lock Cutout, finished":
					return Insert.Dia_lock_Cutout_finished;
				case "Open bottom trash dwr.":
					return Insert.Open_Bot_Trash;
				case "U-Box":
				case "":
				case "None":
					return Insert.No_Insert;
				default:
					return Insert.Unknown;
			}
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
