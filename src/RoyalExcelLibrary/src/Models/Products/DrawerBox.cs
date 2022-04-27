using RoyalExcelLibrary.ExcelUI.Models.Options;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace RoyalExcelLibrary.ExcelUI.Models.Products {

	public enum DBPartType {
		Unknown,
		Side,
		Bottom
	}

	public class DrawerBoxPart : Part {
		public DBPartType PartType { get; set; }
	}

	public enum TrashDrawerType {
		None,
		Single,
		Double,
		DoubleWide
    }

	public class DrawerBox : Product {

		public int JobId { get; set; }
		public double Height { get; set; }
		public double Width { get; set; }
		public double Depth { get; set; }
		public string SideMaterial { get; set; }
		public string BottomMaterial { get; set; }
		public string ClipsOption { get; set; }
		public UndermountNotch NotchOption { get; set; }
		public string InsertOption { get; set; }
		public bool MountingHoles { get; set; }
		public bool ScoopFront { get; set; }
		public bool PullOutFront { get; set; }
		public bool Logo { get; set; }
		public bool LogoInside { get; set; } = true;
		public bool PostFinish { get; set; }
		public TrashDrawerType TrashDrawerType { get; set; } = TrashDrawerType.None;

		public override IEnumerable<Part> GetParts(AppSettings settings) {

			List<Part> parts = new List<Part>();

            DrawerBoxPart front = new DrawerBoxPart {
                PartType = DBPartType.Side,
                Qty = Qty * 2,
                Width = Height,
                Length = Width + settings.ManufacturingValues.FrontBackAdj,
                UseType = InventoryUseType.Linear
            };

			if (Math.Abs(front.Length - 517) < 1) front.Length = 517;

			// If the material is hybrid or economy, and the drawer box has a scoop front, the front of the drawerbox should be solid, while the back will be economy birch
			bool seperateFrontScoop = (SideMaterial == "HybridBirch" || SideMaterial == "BirchFJ") && (ScoopFront || PullOutFront);
			// For double teir cutlery, the back is a different size than the front
			bool doubleTeirCutlery = InsertOption.Equals("Dbl Tier Cutlery");

			if (seperateFrontScoop || doubleTeirCutlery) {

				var backHeight = Height;
				var backLength = Width + settings.ManufacturingValues.FrontBackAdj;
				var backMaterial = "BirchFJ";

				if (doubleTeirCutlery) {
					backHeight = 60;
					backLength = Width - 32;
					backMaterial = SideMaterial;
				}

				DrawerBoxPart back = new DrawerBoxPart {
					PartType = DBPartType.Side,
					Qty = Qty,
					Width = backHeight,
					Length = backLength,
					UseType = InventoryUseType.Linear,
					CutListName = "Back",
					Material = backMaterial
				};

				if (Math.Abs(back.Length - 517) < 1) back.Length = 517;

				front.CutListName = "Front";
				front.Material = "BirchCL";
				front.Qty = Qty;

				parts.Add(front);
				parts.Add(back);
			} else {
				front.CutListName = "Front/Back";
				if (SideMaterial == "Hybrid")
					front.Material = "BirchFJ";
				else front.Material = SideMaterial;
				parts.Add(front);
			}

            DrawerBoxPart sides = new DrawerBoxPart {
                PartType = DBPartType.Side,
                CutListName = "Sides",
                Qty = Qty * 2,
                Width = Height,
                Length = Depth - settings.ManufacturingValues.SideAdj,
                UseType = InventoryUseType.Linear
            };
			if (Math.Abs(sides.Length - 517) < 1) sides.Length = 517;
			if (SideMaterial == "Hybrid")
				sides.Material = "BirchCL";
			else sides.Material = SideMaterial;

			parts.Add(sides);

            DrawerBoxPart bottom = new DrawerBoxPart {
                PartType = DBPartType.Bottom,
                CutListName = "Bottom",
                Width = Width - 2 * settings.ManufacturingValues.SideThickness + 2 * settings.ManufacturingValues.DadoDepth - settings.ManufacturingValues.BottomAdj,
                Length = Depth - 2 * settings.ManufacturingValues.SideThickness + 2 * settings.ManufacturingValues.DadoDepth - settings.ManufacturingValues.BottomAdj,
                Qty = Qty,
                UseType = InventoryUseType.Area,
                Material = BottomMaterial
            };

            parts.Add(bottom);

			var dividerParts = GetDividerParts(settings);
			if (!(dividerParts is null) && dividerParts.Count() > 0)
				parts.AddRange(dividerParts);

			var cutleryParts = GetCutleryParts(settings);
			if (!(cutleryParts is null) && cutleryParts.Count() > 0)
				parts.AddRange(cutleryParts);

			return parts;

		}

		enum CutleryType {
			Unknown,
			CutlerySmall,
			CutleryLarge,
			CutleryTwoTeir
		};

		private IEnumerable<DrawerBoxPart> GetCutleryParts(AppSettings settings) {

			List<DrawerBoxPart> parts = new List<DrawerBoxPart>();

			CutleryType type;
			switch (InsertOption) {
				case "Cutlery 14 5/8":
					type = CutleryType.CutlerySmall;
					break;
				case "Cutlery 20 5/8":
					type = CutleryType.CutleryLarge;
					break;
				case "Dbl Tier Cutlery":
					type = CutleryType.CutleryTwoTeir;
					break;
				default:
					if (InsertOption.ToLower().Contains("cutlery"))
						MessageBox.Show($"Unknown cutlery option found '{InsertOption}'");
					return Enumerable.Empty<DrawerBoxPart>();
            }
			
			if (type == CutleryType.CutleryTwoTeir) {

				double boxHeight = settings.TeiredCutlerySettings.Height;
				double boxWidth = Width - settings.TeiredCutlerySettings.WidthUndersize;
				double boxDepth = Depth - settings.TeiredCutlerySettings.DepthUndersize;

				var levelTwoFrontBack = new DrawerBoxPart() {
					PartType = DBPartType.Side,
					CutListName = "Cutlery Front/Back",
					Qty = Qty * 2,
					Width = boxHeight,
					Length = boxWidth + settings.ManufacturingValues.FrontBackAdj,
					UseType = InventoryUseType.Linear,
					Material = SideMaterial
				};

				parts.Add(levelTwoFrontBack);

				var levelTwoSides = new DrawerBoxPart() {
					PartType = DBPartType.Side,
					CutListName = "Cutlery Sides",
					Qty = Qty * 2,
					Width = boxHeight,
					Length = boxDepth - settings.ManufacturingValues.SideAdj,
					UseType = InventoryUseType.Linear,
					Material = SideMaterial
				};

				parts.Add(levelTwoSides);

				var levelTwoBottoms = new DrawerBoxPart() {
					PartType = DBPartType.Bottom,
					CutListName = "Cutlery Bottom",
					Width = boxWidth,
					Length = boxDepth,
					Qty = Qty,
					UseType = InventoryUseType.Area,
					Material = BottomMaterial
				};

				parts.Add(levelTwoBottoms);

			}


			return parts;

        }

		private IEnumerable<DrawerBoxPart> GetDividerParts(AppSettings settings) {

			List<DrawerBoxPart> parts = new List<DrawerBoxPart>();

			Regex rx = new Regex(@"(?<=Fixed\sDivider\s)[0-9]+", RegexOptions.IgnoreCase);
			MatchCollection matches = rx.Matches(InsertOption);
			if (matches.Count > 0) {

				string strDivCount = matches[0].Value;
				int dividerCount;
				bool read = int.TryParse(strDivCount, out dividerCount);

				if (dividerCount > 0 && read) {

					var dividerHeight = Height;
					if (settings.DividerSettings.Height != -1) {
						dividerHeight = settings.DividerSettings.Height;
                    }

					DrawerBoxPart divider = new DrawerBoxPart {
						PartType = DBPartType.Side,
						CutListName = "Dividers",
						Width = dividerHeight,
						Length = Depth + settings.DividerSettings.LengthAdjustment,
						Qty = dividerCount,
						UseType = InventoryUseType.Linear,
						Material = SideMaterial
					};

					parts.Add(divider);

				}
			}

			return parts;
		}

		public double GetWeight(AppSettings settings) {

			var sizeAdj = 2 * settings.ManufacturingValues.DadoDepth;
			var areaBottom = (Width - sizeAdj) * (Depth - sizeAdj) / 92903; ;

			double bottom_weight = areaBottom;
			if (BottomMaterial.Contains("1/4"))
				bottom_weight *= settings.ManufacturingValues.BottomSqrFtWeight1_4;
			else if (BottomMaterial.Contains("1/2"))
				bottom_weight *= settings.ManufacturingValues.BottomSqrFtWeight1_2;


			var areaSides = (Width * 2 + Depth * 2) * Height / 92903;

			double side_weight = areaSides * settings.ManufacturingValues.SideSqrFtWeight;

			return Qty * (side_weight + bottom_weight);
		}
	}

}
