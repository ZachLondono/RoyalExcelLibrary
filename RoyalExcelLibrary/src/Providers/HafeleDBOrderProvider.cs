using System;
using System.Collections.Generic;
using System.Linq;

using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System.Diagnostics;
using RoyalExcelLibrary.ExportFormat;
using RoyalExcelLibrary.Models.Options;
using ExcelDna.Integration;
using Microsoft.VisualBasic;
using ClosedXML.Excel;

namespace RoyalExcelLibrary.Providers {

	public class HafeleDBOrderProvider : IFileOrderProvider {

		public const decimal MARK_UP = 1.3M;

		public string FilePath { get; set; }

		public Order LoadCurrentOrder() {

			if (string.IsNullOrEmpty(FilePath)) return null;

			var workbook = new XLWorkbook(FilePath);
			var sourceData = workbook.Worksheet("Order Sheet");

			string clientAccountNumber = sourceData.GetStringValue("K5");
			string clientPO = sourceData.GetStringValue("K6");
			string jobName = sourceData.GetStringValue("K7");
			string company = sourceData.GetStringValue("Company");
			Address address = new Address {
				Line1 = sourceData.GetStringValue("V5"),
				Line2 = sourceData.GetStringValue("V6"),
				City = sourceData.GetStringValue("V7"),
				State = sourceData.GetStringValue("V8"),
				Zip = sourceData.GetStringValue("V9")
			};

			decimal grossRevenue = (Decimal.Parse(sourceData.GetStringValue("G13")) - 50M) / 1.3M;
			string hafelePO = sourceData.GetStringValue("K10");
			string hafeleProjectNum = sourceData.GetStringValue("K11");
			string hafeleCfg = "";

			Job job = new Job {
				JobSource = "Hafele",
				Name = jobName,
				GrossRevenue = grossRevenue,
				CreationDate = DateTime.Now
			};

			string sideMaterialStr = sourceData.GetStringValue("Material");
			MaterialType sideMaterial = ParseMaterial(sideMaterialStr);

			bool mountingHoles = sourceData.GetStringValue("MountingHoles").Equals("Yes");
			bool postFinish = sourceData.GetStringValue("PostFinish").Equals("Yes");
			bool setupCharge = sourceData.GetStringValue("LogoOption").Equals("Yes - With Setup");

			IXLCell qtyStart = sourceData.Cell("B16");
			IXLCell heightStart = sourceData.Cell("F16");
			IXLCell widthStart = sourceData.Cell("G16");
			IXLCell depthStart = sourceData.Cell("H16");
			IXLCell scoopStart = sourceData.Cell("I16");
			IXLCell bottomStart = sourceData.Cell("J16");
			IXLCell notchStart = sourceData.Cell("k16");
			IXLCell logoStart = sourceData.Cell("L16");
			IXLCell clipsStart = sourceData.Cell("M16");
			IXLCell accessoryStart = sourceData.Cell("N16");
			IXLCell jobNameStart = sourceData.Cell("O16");
			IXLCell unitPriceStart = sourceData.Cell("P16");
			IXLCell noteStart = sourceData.Cell("S16");
			IXLCell aDimStart = sourceData.Cell("U16");
			IXLCell bDimStart = sourceData.Cell("V16");
			IXLCell cDimStart = sourceData.Cell("W16");

			bool convertToMM = !(sourceData.GetStringValue("Notation").Equals("Metric"));

			List<DrawerBox> boxes = new List<DrawerBox>();

			bool errorsEncountered = false;

			int lineNum = 1;
			int maxCount = 200;
			int i = 0;
			while (i < maxCount) {

				try {

					IXLCell qty = qtyStart.Offset(i, 0);
					if (string.IsNullOrEmpty(qty.GetString()))
						break;

					DrawerBox box;
					if (accessoryStart.Offset(i, 0).GetString().Equals("U-Box")) {
						box = new UDrawerBox();
						(box as UDrawerBox).A = aDimStart.Offset(i, 0).GetDouble() * (convertToMM ? 25.4 : 1);
						(box as UDrawerBox).B = bDimStart.Offset(i, 0).GetDouble() * (convertToMM ? 25.4 : 1);
						(box as UDrawerBox).C = cDimStart.Offset(i, 0).GetDouble() * (convertToMM ? 25.4 : 1);
						box.ProductDescription = "U-Shaped Drawer Box";
					} else {
						box = new DrawerBox {
							ProductDescription = "Strandard Drawer Box"
						};
					}

					box.ProductName = "Drawer Box";
					box.SideMaterial = sideMaterial;
					box.BottomMaterial = ParseMaterial(bottomStart.Offset(i, 0).GetString());
					box.ClipsOption = ParseClips(clipsStart.Offset(i, 0).GetString());
					box.InsertOption = accessoryStart.Offset(i, 0).GetString();
					box.NotchOption = ParseNotch(notchStart.Offset(i, 0).GetString());
					box.Qty = Convert.ToInt32(qty.Offset(i, 0).GetString());
					box.Height = heightStart.Offset(i, 0).GetDouble() * (convertToMM ? 25.4 : 1);
					box.Width = widthStart.Offset(i, 0).GetDouble() * (convertToMM ? 25.4 : 1);
					box.Depth = depthStart.Offset(i, 0).GetDouble() * (convertToMM ? 25.4 : 1);
					box.Logo = logoStart.Offset(i, 0).GetString().Equals("Yes");
					box.ScoopFront = scoopStart.Offset(i, 0).GetString().Equals("Scoop Front");
					box.MountingHoles = mountingHoles;
					box.PostFinish = postFinish;
					box.UnitPrice = Decimal.Parse(unitPriceStart.Offset(i, 0).GetString()) / MARK_UP;
					box.LineNumber = lineNum++;

					box.Note = jobNameStart.Offset(i, 0).GetString();
					box.LevelName = noteStart.Offset(i, 0).GetString();

					boxes.Add(box);

				} catch (Exception e) {
					Debug.WriteLine($"Unable to parse box on line #{i}\n{e}");
					errorsEncountered = true;
				}

				i++;
			}

			if (errorsEncountered)
				System.Windows.Forms.MessageBox.Show("One or more lines could not be parsed due to errors. Please check the order source document.", "Error Loading Order");

			string pronum = Interaction.InputBox("Enter Pro Number", "Pro Number", "none", 0, 0);

			HafeleOrder order = new HafeleOrder(job);
			order.AddProducts(boxes);
			order.Number = hafelePO;
			order.ShippingCost = 0;
			order.Tax = 0;
			order.SubTotal = order.Products.Sum(b => Convert.ToDecimal(b.Qty) * b.UnitPrice) + (setupCharge ? 50 / MARK_UP : 0);
			order.Customer = new Company {
				Name = company,
				Address = address
			};
			order.ConfigNumber = hafeleCfg;
			order.ProjectNumber = hafeleProjectNum;
			order.ProNumber = pronum;
			order.ClientPurchaseOrder = clientPO;
			order.ClientAccountNumber = clientAccountNumber;
			order.SourceFile = FilePath;

			workbook.Dispose();

			return order;

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

	public static class XLExtension {
	
		public static IXLCell Offset(this IXLCell cell, int rows, int columns) {
			var address = cell.Address;
			var worksheet = cell.Worksheet;
			return worksheet.Cell(address.RowNumber + rows, address.ColumnNumber + columns);
        }

		public static string GetStringValue(this IXLWorksheet worksheet, string range) {
			var cell = worksheet.Cell(range);
			if (cell is null) return "";
			return cell.GetString();
		}

	}

}
