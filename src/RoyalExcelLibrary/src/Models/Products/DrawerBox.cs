using RoyalExcelLibrary.ExcelUI.Models.Options;
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
		public MaterialType SideMaterial { get; set; }
		public MaterialType BottomMaterial { get; set; }
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

            if ((SideMaterial == MaterialType.HybridBirch || SideMaterial == MaterialType.EconomyBirch) && (ScoopFront || PullOutFront)) {
				// If the material is hybrid or economy, and the drawer box has a scoop front, the front of the drawerbox should be solid, while the back will be economy birch
                DrawerBoxPart back = new DrawerBoxPart {
                    PartType = DBPartType.Side,
                    Qty = Qty,
                    Width = Height,
                    Length = Width + settings.ManufacturingValues.FrontBackAdj,
                    UseType = InventoryUseType.Linear,
                    CutListName = "Back",
                    Material = MaterialType.EconomyBirch
                };

                front.CutListName = "Front";
				front.Material = MaterialType.SolidBirch;
				front.Qty = Qty;

				parts.Add(front);
				parts.Add(back);
			} else {
				front.CutListName = "Front/Back";
				if (SideMaterial == MaterialType.HybridBirch)
					front.Material = MaterialType.EconomyBirch;
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
            if (SideMaterial == MaterialType.HybridBirch)
				sides.Material = MaterialType.SolidBirch;
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

				var levelTwoFrontBack = new DrawerBoxPart() {
					PartType = DBPartType.Side,
					CutListName = "Cutlery Front/Back",
					Qty = Qty * 2,
					Width = settings.TeiredCutlerySettings.Height,
					Length = (Width - settings.TeiredCutlerySettings.WidthUndersize) + settings.ManufacturingValues.FrontBackAdj,
					UseType = InventoryUseType.Linear
				};

				parts.Add(levelTwoFrontBack);

				var levelTwoSides = new DrawerBoxPart() {
					PartType = DBPartType.Side,
					CutListName = "Cutlery Sides",
					Qty = Qty * 2,
					Width = settings.TeiredCutlerySettings.Height,
					Length = (Depth - settings.TeiredCutlerySettings.DepthUndersize) - settings.ManufacturingValues.SideAdj,
					UseType = InventoryUseType.Linear
				};

				parts.Add(levelTwoSides);

				var levelTwoBottoms = new DrawerBoxPart() {
					PartType = DBPartType.Bottom,
					CutListName = "Cutlery Bottom",
					Width = (Width - settings.TeiredCutlerySettings.WidthUndersize) - 2 * settings.ManufacturingValues.SideThickness + 2 * settings.ManufacturingValues.DadoDepth - settings.ManufacturingValues.BottomAdj,
					Length = (Depth - settings.TeiredCutlerySettings.DepthUndersize) - 2 * settings.ManufacturingValues.SideThickness + 2 * settings.ManufacturingValues.DadoDepth - settings.ManufacturingValues.BottomAdj,
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
			if (BottomMaterial is MaterialType.BlackMela1_4 || BottomMaterial is MaterialType.WhiteMela1_4 || BottomMaterial is MaterialType.Plywood1_4)
				bottom_weight *= settings.ManufacturingValues.BottomSqrFtWeight1_4;
			else if (BottomMaterial is MaterialType.BlackMela1_2 || BottomMaterial is MaterialType.WhiteMela1_2 || BottomMaterial is MaterialType.Plywood1_2)
				bottom_weight *= settings.ManufacturingValues.BottomSqrFtWeight1_2;


			var areaSides = (Width * 2 + Depth * 2) * Height / 92903;

			double side_weight = areaSides * settings.ManufacturingValues.SideSqrFtWeight;

			return Qty * (side_weight + bottom_weight);
		}
	}

}
