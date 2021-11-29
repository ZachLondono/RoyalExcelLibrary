using RoyalExcelLibrary.Models.Options;
using System.Collections.Generic;

namespace RoyalExcelLibrary.Models.Products {

	public enum DBPartType {
		Unknown,
		Side,
		Bottom
	}

	public class DrawerBoxPart : Part {
		public DBPartType PartType { get; set; }
	}

	public class DrawerBox : Product {
		public int JobId { get; set; }
		public double Height { get; set; }
		public double Width { get; set; }
		public double Depth { get; set; }
		public MaterialType SideMaterial { get; set; }
		public MaterialType BottomMaterial { get; set; }
		public Clips ClipsOption { get; set; }
		public UndermountNotch NotchOption { get; set; }
		public string InsertOption { get; set; }
		public bool MountingHoles { get; set; }
		public bool ScoopFront { get; set; }
		public bool Logo { get; set; }
		public bool PostFinish { get; set; }

		public override IEnumerable<Part> GetParts() {

			List<Part> parts = new List<Part>();

            DrawerBoxPart front = new DrawerBoxPart {
                PartType = DBPartType.Side,
                Qty = Qty * 2,
                Width = Height,
                Length = Width + ManufacturingConstants.FrontBackAdj,
                UseType = InventoryUseType.Linear
            };

            if ((SideMaterial == MaterialType.HybridBirch || SideMaterial == MaterialType.EconomyBirch) && ScoopFront) {
                DrawerBoxPart back = new DrawerBoxPart {
                    PartType = DBPartType.Side,
                    Qty = Qty * 2,
                    Width = Height,
                    Length = Width + ManufacturingConstants.FrontBackAdj,
                    UseType = InventoryUseType.Linear,
                    CutListName = "Back",
                    Material = MaterialType.EconomyBirch
                };

                front.CutListName = "Front";
				front.Material = MaterialType.SolidBirch;

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
                Length = Depth - ManufacturingConstants.SideAdj,
                UseType = InventoryUseType.Linear
            };
            if (SideMaterial == MaterialType.HybridBirch)
				sides.Material = MaterialType.SolidBirch;
			else sides.Material = SideMaterial;

			parts.Add(sides);

            DrawerBoxPart bottom = new DrawerBoxPart {
                PartType = DBPartType.Bottom,
                CutListName = "Bottom",
                Width = Width - 2 * ManufacturingConstants.SideThickness + 2 * ManufacturingConstants.DadoDepth - ManufacturingConstants.BottomAdj,
                Length = Depth - 2 * ManufacturingConstants.SideThickness + 2 * ManufacturingConstants.DadoDepth - ManufacturingConstants.BottomAdj,
                Qty = Qty,
                UseType = InventoryUseType.Area,
                Material = BottomMaterial
            };

            parts.Add(bottom);

			return parts;

		}

		public double Weight {
			get {

				var sizeAdj = 2 * ManufacturingConstants.DadoDepth;
				var areaBottom = (Width - sizeAdj) * (Depth - sizeAdj) / 92903; ;

				double bottom_weight = areaBottom;
				if (BottomMaterial is MaterialType.BlackMela1_4 || BottomMaterial is MaterialType.WhiteMela1_4 || BottomMaterial is MaterialType.Plywood1_4)
					bottom_weight *= ManufacturingConstants.BottomSqrFtWeight1_4;
				else if (BottomMaterial is MaterialType.BlackMela1_2 || BottomMaterial is MaterialType.WhiteMela1_2 || BottomMaterial is MaterialType.Plywood1_2)
					bottom_weight *= ManufacturingConstants.BottomSqrFtWeight1_2;


				var areaSides = (Width * 2 + Depth * 2) * Height / 92903;

				double side_weight = areaSides * ManufacturingConstants.SideSqrFtWeight;

				return Qty * (side_weight + bottom_weight);

			}
		}
	}

}
