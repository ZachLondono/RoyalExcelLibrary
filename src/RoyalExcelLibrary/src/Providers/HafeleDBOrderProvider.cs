using System;
using System.Collections.Generic;
using System.Linq;

using RoyalExcelLibrary.ExcelUI.Models;
using RoyalExcelLibrary.ExcelUI.Models.Products;
using System.Diagnostics;
using RoyalExcelLibrary.ExcelUI.ExportFormat;
using RoyalExcelLibrary.ExcelUI.Models.Options;
using ExcelDna.Integration;
using Microsoft.VisualBasic;
using ClosedXML.Excel;

namespace RoyalExcelLibrary.ExcelUI.Providers {

	public class HafeleDBOrderProvider : IFileOrderProvider {

		public string FilePath { get; set; }

		public Order LoadCurrentOrder() {

			if (string.IsNullOrEmpty(FilePath)) return null;

			Order order;

			using (var workbook = new XLWorkbook(FilePath)) {

				int version = GetHafeleVersionNum(workbook);

				switch (version) {
					case 2:
						order = LoadV2Order(workbook);
						break;
					case 3:
						order = LoadV3Order(workbook);
						break;
					case 1:
					default:
						order = LoadV1Order(workbook);
						break;
				}
			}

			return order;
		
		}

		private Order LoadV1Order(XLWorkbook workbook) {
			return null;
        }

		private Order LoadV2Order(XLWorkbook workbook) {
			
			Data data = new Data();
			var sourceData = workbook.Worksheet("Order Sheet");
			

			data.clientAccountNumber = sourceData.GetStringValue("K5");
			data.clientPO = sourceData.GetStringValue("K6");
			data.jobName = sourceData.GetStringValue("K7");
			data.company = sourceData.GetStringValue("Company");
			data.address = new Address {
				Line1 = sourceData.GetStringValue("V5").Trim(),
				Line2 = sourceData.GetStringValue("V6").Trim(),
				City = sourceData.GetStringValue("V7").Trim(),
				State = sourceData.GetStringValue("V8").Trim(),
				Zip = sourceData.GetStringValue("V9").Trim()
			};

			try {
				decimal markup = (decimal)workbook.Range("StdMarkup").FirstCell().GetDoubleValue();
				data.markup = markup;
			} catch {
				data.markup = 1.3M;
			}

			var delivered = sourceData.GetStringValue("G13").Replace("$", String.Empty);
			data.grossRevenue = string.IsNullOrEmpty(delivered) ? 0 : (Decimal.Parse(delivered) - 50M) / data.markup;
			data.hafelePO = sourceData.GetStringValue("K10");
			data.hafeleProjectNum = sourceData.GetStringValue("K11");

			data.qtyStart = sourceData.Cell("B16");
			data.heightStart = sourceData.Cell("F16");
			data.widthStart = sourceData.Cell("G16");
			data.depthStart = sourceData.Cell("H16");
			data.scoopStart = sourceData.Cell("I16");
			data.bottomStart = sourceData.Cell("J16");
			data.notchStart = sourceData.Cell("k16");
			data.logoStart = sourceData.Cell("L16");
			data.clipsStart = sourceData.Cell("M16");
			data.accessoryStart = sourceData.Cell("N16");
			data.jobNameStart = sourceData.Cell("O16");
			data.unitPriceStart = sourceData.Cell("P16");
			data.noteStart = sourceData.Cell("S16");
			data.aDimStart = sourceData.Cell("U16");
			data.bDimStart = sourceData.Cell("V16");
			data.cDimStart = sourceData.Cell("W16");

			string sideMaterialStr = sourceData.GetStringValue("Material");
			data.sideMaterial = ParseMaterial(sideMaterialStr);
			data.mountingHoles = sourceData.GetStringValue("MountingHoles").Equals("Yes");
			data.postFinish = sourceData.GetStringValue("PostFinish").Equals("Yes");
			data.setupCharge = sourceData.GetStringValue("LogoOption").Equals("Yes - With Setup");
			data.convertToMM = !(sourceData.GetStringValue("Notation").Equals("Metric"));

			return LoadOrderHelper(data);

		}

		private Order LoadV3Order(XLWorkbook workbook) {
			var sourceData = workbook.Worksheet("Order Sheet");
			Data data = new Data();

			data.clientAccountNumber = sourceData.GetStringValue("K6");
			data.clientPO = sourceData.GetStringValue("K7");
			data.jobName = sourceData.GetStringValue("K8");
			data.company = sourceData.GetStringValue("Company");
			data.address = new Address {
				Line1 = sourceData.GetStringValue("V6").Trim(),
				Line2 = sourceData.GetStringValue("V7").Trim(),
				City = sourceData.GetStringValue("V8").Trim(),
				State = sourceData.GetStringValue("V9").Trim(),
				Zip = sourceData.GetStringValue("V10").Trim()
			};

			try {
				decimal markup = (decimal)workbook.Range("StdMarkup").FirstCell().GetDoubleValue();
				data.markup = markup;
			} catch {
				data.markup = 1.3M;
            }

			data.grossRevenue = (Decimal.Parse(sourceData.GetStringValue("G14")) - 50M) / data.markup;
			data.hafelePO = sourceData.GetStringValue("K11");
			data.hafeleProjectNum = sourceData.GetStringValue("K12");

			data.qtyStart = sourceData.Cell("B17");
			data.heightStart = sourceData.Cell("F17");
			data.widthStart = sourceData.Cell("G17");
			data.depthStart = sourceData.Cell("H17");
			data.scoopStart = sourceData.Cell("I17");
			data.bottomStart = sourceData.Cell("J17");
			data.notchStart = sourceData.Cell("k17");
			data.logoStart = sourceData.Cell("L17");
			data.clipsStart = sourceData.Cell("M17");
			data.accessoryStart = sourceData.Cell("N17");
			data.jobNameStart = sourceData.Cell("O17");
			data.unitPriceStart = sourceData.Cell("P17");
			data.noteStart = sourceData.Cell("S17");
			data.aDimStart = sourceData.Cell("U17");
			data.bDimStart = sourceData.Cell("V17");
			data.cDimStart = sourceData.Cell("W17");

			string sideMaterialStr = sourceData.GetStringValue("Material");
			data.sideMaterial = ParseMaterial(sideMaterialStr);
			data.mountingHoles = sourceData.GetStringValue("MountingHoles").Equals("Yes");
			data.postFinish = sourceData.GetStringValue("PostFinish").Equals("Yes");
			data.setupCharge = sourceData.GetStringValue("LogoOption").Equals("Yes - With Setup");
			data.convertToMM = !(sourceData.GetStringValue("Notation").Equals("Metric"));

			return LoadOrderHelper(data);

		}

		private Order LoadOrderHelper(Data data) { 
			
			string hafeleCfg = "";
			Job job = new Job {
				JobSource = "Hafele",
				Name = data.jobName,
				GrossRevenue = data.grossRevenue,
				CreationDate = DateTime.Now
			};

			List<DrawerBox> boxes = new List<DrawerBox>();

			int lineNum = 1;
			int maxCount = 200;
			int i = 0;
			while (i < maxCount) {

				try {

					IXLCell qty = data.qtyStart.Offset(i, 0);

					string qtyStr = qty.GetStringValue();
					if (string.IsNullOrEmpty(qtyStr)) {
						i++;
						continue;
					} else if (i > 200 || qtyStr.Equals("End")) break;

					DrawerBox box;
					if (data.accessoryStart.Offset(i, 0).GetStringValue().Equals("U-Box")) {
						box = new UDrawerBox();
						(box as UDrawerBox).A = data.aDimStart.Offset(i, 0).GetDoubleValue() * (data.convertToMM ? 25.4 : 1);
						(box as UDrawerBox).B = data.bDimStart.Offset(i, 0).GetDoubleValue() * (data.convertToMM ? 25.4 : 1);
						(box as UDrawerBox).C = data.cDimStart.Offset(i, 0).GetDoubleValue() * (data.convertToMM ? 25.4 : 1);
						box.ProductDescription = "U-Shaped Drawer Box";
					} else {
						box = new DrawerBox {
							ProductDescription = "Strandard Drawer Box"
						};
					}

					box.ProductName = "Drawer Box";
					box.SideMaterial = data.sideMaterial;
					box.BottomMaterial = ParseMaterial(data.bottomStart.Offset(i, 0).GetStringValue());
					box.ClipsOption = ParseClips(data.clipsStart.Offset(i, 0).GetStringValue());
					box.InsertOption = data.accessoryStart.Offset(i, 0).GetStringValue();
					box.NotchOption = ParseNotch(data.notchStart.Offset(i, 0).GetStringValue());

					box.Qty = string.IsNullOrEmpty(qtyStr) ? 0 : Convert.ToInt32(qtyStr);
					box.Height = data.heightStart.Offset(i, 0).GetDoubleValue() * (data.convertToMM ? 25.4 : 1);
					box.Width = data.widthStart.Offset(i, 0).GetDoubleValue() * (data.convertToMM ? 25.4 : 1);
					box.Depth = data.depthStart.Offset(i, 0).GetDoubleValue() * (data.convertToMM ? 25.4 : 1);
					box.Logo = data.logoStart.Offset(i, 0).GetStringValue().Equals("Yes");
					box.ScoopFront = data.scoopStart.Offset(i, 0).GetStringValue().Equals("Scoop Front");
					box.PostFinish = data.postFinish;
					box.MountingHoles = data.mountingHoles;
					
					string unitPriceStr = data.unitPriceStart.Offset(i, 0).GetStringValue();
					box.UnitPrice = string.IsNullOrEmpty(unitPriceStr) ? 0 : Decimal.Parse(unitPriceStr) / data.markup;
					box.LineNumber = lineNum++;

					box.LevelName = data.jobNameStart.Offset(i, 0).GetStringValue();
					box.Note = data.noteStart.Offset(i, 0).GetStringValue();

					boxes.Add(box);

				} catch (Exception e) {
					Debug.WriteLine($"Unable to parse box on line #{i}\n{e}");
				}

				i++;
			}

			HafeleOrder order = new HafeleOrder(job);
			order.AddProducts(boxes);
			order.Number = data.hafelePO;
			order.ShippingCost = 0;
			order.Tax = 0;
			order.SubTotal = order.Products.Sum(b => Convert.ToDecimal(b.Qty) * b.UnitPrice) + (data.setupCharge ? 50 / data.markup : 0);
			order.Customer = new Company {
				Name = data.company,
				Address = data.address
			};
			order.ConfigNumber = hafeleCfg;
			order.ProjectNumber = data.hafeleProjectNum;
			order.ClientPurchaseOrder = data.clientPO;
			order.ClientAccountNumber = data.clientAccountNumber;
			order.SourceFile = FilePath;

			return order;

		}

		private int GetHafeleVersionNum(XLWorkbook workbook) {

			IXLWorksheet dataSheet;
			bool getSheet = workbook.Worksheets.TryGetWorksheet("Data", out dataSheet);
			if (getSheet) {

				try {
					var range = dataSheet.Range("MajorVersion");
					return int.Parse(range.FirstCell().GetString());
                } catch {
					return -1;
                }

            }

			return -1;

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
				case "Walnut, finished":
					return MaterialType.Walnut;
				case "Walnut, unfinished":
					return MaterialType.UnFinishedWalnut;
				case "White Oak":
				case "White Oak, finished":
					return MaterialType.WhiteOak;
				case "White Oak, unfinished":
					return MaterialType.UnFinishedWhiteOak;
				case "1/4\" Plywood":
					return MaterialType.Plywood1_4;
				case "1/2\" Plywood":
					return MaterialType.Plywood1_2;
				default:
					return MaterialType.Unknown;
			}

		}

		struct Data {
			public string company {get; set;}
			public MaterialType sideMaterial {get; set;}
			public bool mountingHoles {get; set;}
			public bool postFinish {get; set;}
			public bool setupCharge {get; set;}
			public bool convertToMM {get; set;}
			public string clientAccountNumber {get; set;}
			public string clientPO {get; set;}
			public string jobName {get; set;}
			public Address address {get; set;}
			public decimal grossRevenue {get; set;}
			public string hafelePO {get; set;}
			public string hafeleProjectNum {get; set;}
			public IXLCell qtyStart {get; set;}
			public IXLCell heightStart {get; set;}
			public IXLCell widthStart {get; set;}
			public IXLCell depthStart {get; set;}
			public IXLCell scoopStart {get; set;}
			public IXLCell bottomStart {get; set;}
			public IXLCell notchStart {get; set;}
			public IXLCell logoStart {get; set;}
			public IXLCell clipsStart {get; set;}
			public IXLCell accessoryStart {get; set;}
			public IXLCell jobNameStart {get; set;}
			public IXLCell unitPriceStart {get; set;}
			public IXLCell noteStart {get; set;}
			public IXLCell aDimStart {get; set;}
			public IXLCell bDimStart {get; set;}
			public IXLCell cDimStart { get; set; }
			public decimal markup { get; set; }
		}

	}

	public static class XLExtension {

		public static IXLCell Offset(this IXLCell cell, int rows, int columns) {
			var address = cell.Address;
			var worksheet = cell.Worksheet;
			return worksheet.Cell(address.RowNumber + rows, address.ColumnNumber + columns);
		}

		// Returns the value of the cell as a String
		public static string GetStringValue(this IXLCell cell) {

			if (cell.HasFormula) {
				return cell.CachedValue.ToString();
			}

			return cell.RichText.ToString();
			/*
						string val;
						if (cell.TryGetValue<string>(out val)) return val;
						return "";*/

		}

		public static double GetDoubleValue(this IXLCell cell) {

			string value;

			if (cell.HasFormula) {
				value = cell.CachedValue.ToString();
			} else {
				value = cell.RichText.ToString();
			}

			return HelperFuncs.ConvertToDouble(value);


		}

		public static string GetStringValue(this IXLWorksheet worksheet, string range) {
			var cell = worksheet.Cell(range);
			if (cell is null) return "";
			return cell.GetStringValue();
		}

	}

}
