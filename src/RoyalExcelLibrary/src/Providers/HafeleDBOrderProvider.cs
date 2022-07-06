using System;
using System.Collections.Generic;
using System.Linq;

using RoyalExcelLibrary.ExcelUI.Models;
using RoyalExcelLibrary.ExcelUI.Models.Products;
using System.Diagnostics;
using RoyalExcelLibrary.ExcelUI.ExportFormat;
using RoyalExcelLibrary.ExcelUI.Models.Options;
using ExcelDna.Integration;
using ClosedXML.Excel;
using System.Windows.Forms;
using RoyalExcelLibrary.ExcelUI.src.FluentWorkbookValidation;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace RoyalExcelLibrary.ExcelUI.Providers {

    public class HafeleDBOrderProvider : IFileOrderProvider {

		public string FilePath { get; set; }

		private bool unknownLogoFound = false;
		private bool unknownMaterialFound = false;
		private bool unknownNotchFound = false;
		private bool unknownScoopFount = false;

		private readonly AppSettings _settings;

		public HafeleDBOrderProvider() {
			_settings = HelperFuncs.ReadSettings();
		}

		public Order LoadCurrentOrder() {

			if (string.IsNullOrEmpty(FilePath)) return null;

			Order order;

			// If the file is an ods file, convert it to an excel file
			if (Path.GetExtension(FilePath) == ".ods") {

				try {

					var app = ExcelDnaUtil.Application as Excel.Application;
					var wb = app.Workbooks.Open(FilePath, ReadOnly:true);

				
					wb.Worksheets["Order_Sheet"].Name = "Order Sheet";
					string newFilePath = Path.ChangeExtension(FilePath, "xlsx");

					wb.SaveAs(newFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook);
					wb.Close();
	
					FilePath = newFilePath;

				} catch (Exception ex) {
					var response = MessageBox.Show("An error occurred while attempting to convert .ods file to .xlsx file, manual conversion required\nClick 'OK' to show details.", "Conversion Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
					if (response == DialogResult.OK) {
						MessageBox.Show(ex.ToString(), "Error Details");
                    }
				}

				MessageBox.Show("Order file was converted from .ods to .xlsx\nMake sure to validate data integrity.", "File Conversion", MessageBoxButtons.OK, MessageBoxIcon.Information);

			}

			using (var workbook = new XLWorkbook(FilePath)) {

				int version = GetHafeleVersionNum(workbook);

				if (version != 3)
					MessageBox.Show("Old workbook version, manual verification required", "Version Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

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
			throw new NotImplementedException("Workbook version to old, not supported");
        }

		private Order LoadV2Order(XLWorkbook workbook) {

			Data data = new Data();
			var sourceData = workbook.Worksheet("Order Sheet");

			try {
				data.OrderDate = DateTime.Parse(sourceData.GetStringValue("OrderDate"));
			} catch {
				data.OrderDate = DateTime.Today;
			}

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

			data.BoxCount = Convert.ToInt32(sourceData.Range("G12").FirstCell().GetDoubleValue());

			try {
				decimal markup = (decimal)workbook.Range("StdMarkup").FirstCell().GetDoubleValue();
				data.markup = markup;
			} catch {
				data.markup = 1.3M;
			}

			var delivered = sourceData.GetStringValue("G13").Replace("$", string.Empty);
			data.grossRevenue = string.IsNullOrEmpty(delivered) ? 0 : (decimal.Parse(delivered) - 50M) / data.markup;
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
			data.sideMaterial = ParseMaterial(sideMaterialStr, out bool postFinish);
			data.mountingHoles = sourceData.GetStringValue("MountingHoles").Equals("Yes");
			data.postFinish = sourceData.GetStringValue("PostFinish").Equals("Yes") || postFinish;
			data.setupCharge = sourceData.GetStringValue("LogoOption").Equals("Yes - With Setup");
			data.convertToMM = !(sourceData.GetStringValue("Notation").Equals("Metric"));

			return LoadOrderHelper(data);

		}

		private Order LoadV3Order(XLWorkbook workbook) {

			WkbkValidator validator = new WkbkValidator(workbook);

			validator.WkbkRule()
						.HasSheet("Order Sheet")
						.WithMessage("Workbook does not contain sheet named 'Order Sheet'");

			validator.WkbkRule()
						.ForSheet("Order Sheet")
						.HasRange("Company")
						.WithMessage("'Order Sheet' is missing named range 'Company'");

			validator.WkbkRule()
						.ForSheet("Order Sheet")
						.ForRange("K6")
						.NotEmpty()
						.WithMessage("Client account number cannot be empty");

			validator.WkbkRule()
						.ForSheet("Order Sheet")
						.ForRange("K7")
						.NotEmpty()
						.WithMessage("Purchase Order field cannot be empty");

			validator.WkbkRule()
						.ForRange("StdMarkup")
						.NotEmpty()
						.ContainsDouble()
						.WithMessage("Unable to read price markup from range 'StdMarkup'");

			validator.WkbkRule()
						.ForSheet("Order Sheet")
						.HasRange("Material")
						.WithMessage("Material field cannot be empty");

			validator.WkbkRule()
						.ForRange("Material")
						.NotEmpty()
						.WithMessage("Material field cannot be empty");


			validator.WkbkRule()
						.ForSheet("Order Sheet")
						.HasRange("BotThickness")
						.WithMessage("Material field cannot be empty");

			validator.WkbkRule()
						.ForRange("BotThickness")
						.NotEmpty()
						.WithMessage("Material field cannot be empty");

			validator.WkbkRule()
						.ForSheet("Order Sheet")
						.HasRange("MountingHoles")
						.WithMessage("Named range 'MountingHoles' cannot be found");

			validator.WkbkRule()
						.ForSheet("Order Sheet")
						.HasRange("PostFinish")
						.WithMessage("Named range 'PostFinish' cannot be found");

			validator.WkbkRule()
						.ForSheet("Order Sheet")
						.HasRange("LogoOption")
						.WithMessage("Named range 'LogoOption' cannot be found");

			validator.WkbkRule()
						.ForSheet("Order Sheet")
						.HasRange("Notation")
						.WithMessage("Named range 'Notation' cannot be found");

			validator.Validate();

			var sourceData = workbook.Worksheet("Order Sheet");
			Data data = new Data();

			try { 
				data.OrderDate = DateTime.Parse(sourceData.GetStringValue("OrderDate"));
			} catch {
				data.OrderDate = DateTime.Today;
			}

			data.OrderNote = sourceData.GetStringValue("N12");
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

			data.BoxCount = Convert.ToInt32(sourceData.Range("G13").FirstCell().GetDoubleValue());

			try {
				decimal markup = (decimal)workbook.Range("StdMarkup").FirstCell().GetDoubleValue();
				data.markup = markup;
			} catch {
				data.markup = 1.3M;
            }

            decimal shippingCost;
            try {
				if (sourceData.GetStringValue("DeliverySelection").Equals("Standard Pallet"))
					shippingCost = 25M;
				else
					shippingCost = 0M;
			} catch {
				shippingCost = 25M;
				MessageBox.Show("Unable to read order shipping method", "Shipping Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}

			decimal grossRevenue = 0;
			try {
				grossRevenue = (decimal.Parse(sourceData.GetStringValue("G14")) - shippingCost) / data.markup;
			} catch {
				MessageBox.Show("Unable to read order price", "Price Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}

			data.grossRevenue = grossRevenue;
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
            data.sideMaterial = ParseMaterial(sideMaterialStr, out bool postFinish);
            data.mountingHoles = sourceData.GetStringValue("MountingHoles").Equals("Yes");
			data.postFinish = sourceData.GetStringValue("PostFinish").Equals("Yes") || postFinish;
			data.convertToMM = !sourceData.GetStringValue("Notation").Equals("Metric");

			data.setupCharge = false;
			data.logoInside = true;
			string logoOptionValue = sourceData.GetStringValue("LogoOption");
			(data.setupCharge, data.logoInside) = ParseLogoOption(logoOptionValue);

			return LoadOrderHelper(data);

		}

		private Order LoadOrderHelper(Data data) { 
			
			string hafeleCfg = "";
			Job job = new Job {
				JobSource = "Hafele",
				Name = data.jobName,
				GrossRevenue = data.grossRevenue,
				CreationDate = data.OrderDate
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
					bool containsUDim = !string.IsNullOrEmpty(data.aDimStart.Offset(i, 0).GetStringValue())
										|| !string.IsNullOrEmpty(data.bDimStart.Offset(i, 0).GetStringValue())
										|| !string.IsNullOrEmpty(data.cDimStart.Offset(i, 0).GetStringValue())
										|| data.accessoryStart.Offset(i, 0).GetStringValue().ToLower().Equals("u-box");

					if (containsUDim) {
						box = new UDrawerBox();

						if (string.IsNullOrEmpty(data.aDimStart.Offset(i, 0).GetStringValue()))
							MessageBox.Show("Missing UBox 'A' Dimension", "Missing Dimension", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        else (box as UDrawerBox).A = data.aDimStart.Offset(i, 0).GetDoubleValue() * (data.convertToMM ? 25.4 : 1);

						if (string.IsNullOrEmpty(data.bDimStart.Offset(i, 0).GetStringValue()))
							MessageBox.Show("Missing UBox 'B' Dimension", "Missing Dimension", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						else (box as UDrawerBox).B = data.bDimStart.Offset(i, 0).GetDoubleValue() * (data.convertToMM ? 25.4 : 1);

						if (string.IsNullOrEmpty(data.cDimStart.Offset(i, 0).GetStringValue()))
							MessageBox.Show("Missing UBox 'C' Dimension", "Missing Dimension", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						else (box as UDrawerBox).C = data.cDimStart.Offset(i, 0).GetDoubleValue() * (data.convertToMM ? 25.4 : 1);

						box.ProductDescription = "U-Shaped Drawer Box";
					} else {
						box = new DrawerBox {
							ProductDescription = "Standard Drawer Box"
						};
					}

					box.ProductName = "Drawer Box";
					box.SideMaterial = data.sideMaterial;
					box.BottomMaterial = ParseMaterial(data.bottomStart.Offset(i, 0).GetStringValue(), out _);
					box.ClipsOption = data.clipsStart.Offset(i, 0).GetStringValue();
					box.NotchOption = ParseNotch(data.notchStart.Offset(i, 0).GetStringValue());
					box.InsertOption = data.accessoryStart.Offset(i, 0).GetStringValue();
					box.TrashDrawerType = ParseTrashType(box.InsertOption);

					box.Qty = string.IsNullOrEmpty(qtyStr) ? 0 : Convert.ToInt32(qtyStr);
					box.Height = data.heightStart.Offset(i, 0).GetDoubleValue() * (data.convertToMM ? 25.4 : 1);
					box.Width = data.widthStart.Offset(i, 0).GetDoubleValue() * (data.convertToMM ? 25.4 : 1);
					box.Depth = data.depthStart.Offset(i, 0).GetDoubleValue() * (data.convertToMM ? 25.4 : 1);
					
					string logoOption = data.logoStart.Offset(i, 0).GetStringValue();
					switch (logoOption.ToLower()) {
						case "yes":
						case "logo":
							box.Logo = true;
							break;
						case "no":
						case "":
						case null:
							box.Logo = false;
							break;
						default:
							unknownLogoFound = true;
							break;
					}
					box.LogoInside = data.logoInside;
					box.PostFinish = data.postFinish;
					box.MountingHoles = data.mountingHoles;

					string scoopOption = data.scoopStart.Offset(i, 0).GetStringValue();
					switch (scoopOption.ToLower()) {
						case "scoop front":
						case "yes":
							box.ScoopFront = true;
							break;
						case "":
						case null:
						case "no":
							box.ScoopFront = false;
							break;
						default:
							unknownScoopFount = true;
							break;
					}
					


					string unitPriceStr = data.unitPriceStart.Offset(i, 0).GetStringValue();
					box.UnitPrice = string.IsNullOrEmpty(unitPriceStr) ? 0 : Decimal.Parse(unitPriceStr) / data.markup;
					box.LineNumber = lineNum++;

					box.LevelName = data.jobNameStart.Offset(i, 0).GetStringValue();
					box.Note = data.noteStart.Offset(i, 0).GetStringValue();

					boxes.Add(box);

				} catch (Exception e) {
					Debug.WriteLine($"Unable to parse box on line #{i + 1}\n{e}");
					var result = MessageBox.Show($"Error encounterd when reading box on line #{i+1}\nClick 'Ok' to skip or 'Cancel' to stop order input", "Parse Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
					if (result == DialogResult.Cancel) {
						return null;
                    }
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

			order.Comment = data.OrderNote;
			if (!string.IsNullOrEmpty(data.OrderNote)) {
				MessageBox.Show(data.OrderNote, "Order Comment");
            }

			if (data.BoxCount != order.Products.Sum(p => p.Qty)) MessageBox.Show("Box count read does not match order data", "Item Count Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

			if (unknownLogoFound) MessageBox.Show("Unknown LOGO option found, manual verification required", "Logo Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			if (unknownMaterialFound) MessageBox.Show("Unknown MATERIAL option found, manual verification required", "Material Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			if (unknownNotchFound) MessageBox.Show("Unknown NOTCH option found, manual verification required", "Notch Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			if (unknownScoopFount) MessageBox.Show("Unknown SCOOP FRONT option found, manual verification required", "Scoop Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

			return order;

		}

		private (bool setupCharge, bool logoInside) ParseLogoOption(string logoOption) {

			bool setupCharge = false;
			bool logoInside = true;

			switch (logoOption) {
				case "Yes-Inside w/ Setup":
				case "Yes - With Setup":
					setupCharge = true;
					logoInside = true;
					break;
				case "Yes":
				case "Yes-Inside":
					setupCharge = false;
					logoInside = true;
					break;
				case "Yes-Outside w/ Setup":
					setupCharge = true;
					logoInside = false;
					break;
				case "Yes-Outside":
					setupCharge = false;
					logoInside = false;
					break;
				case null:
				case "":
				case "No":
					setupCharge = false;
					logoInside = true;
					break;
				default:
					unknownLogoFound = true;
					break;
			}

			return (setupCharge, logoInside);
		}

        private TrashDrawerType ParseTrashType(string insertOption) {
            
			switch (insertOption) {

				case "Trash Drw. Single":
				case "Trash Drw. Single w/ Can":
					return TrashDrawerType.Single;
				case "Trash Drw. Double":
				case "Trash Drw. Double w/ Cans":
					return TrashDrawerType.Double;
				case "Trash Drw. Double Wide":
				case "Trash Drw. Dbl Wide w/ Cans":
					return TrashDrawerType.DoubleWide;
				default:
					return TrashDrawerType.None;
            }

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
					unknownNotchFound = true;
					return UndermountNotch.Unknown;
			}
		}

		private string ParseMaterial(string name, out bool postfinish) {
			postfinish = false;

			if (name.ToLower().Contains("unfinished")) postfinish = false;
			else if (name.ToLower().Contains("finished")) postfinish = true;

			var profile = _settings.MaterialProfiles["hafele"];
			return profile[name];

		}

		struct Data {
			public DateTime OrderDate { get; set; }
			public int BoxCount { get; set; }
			public string OrderNote { get; set; }
			public string company {get; set;}
			public string sideMaterial {get; set;}
			public bool mountingHoles {get; set;}
			public bool postFinish {get; set;}
			public bool setupCharge {get; set;}
			public bool logoInside { get; set; }
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

}
