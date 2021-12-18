using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.ExcelUI.Models.Products {
	public class UDrawerBox : DrawerBox {
		
		public Double A { get; set; }
		public Double B { get; set; }
		public Double C { get; set; }

		public override IEnumerable<Part> GetParts() {

			List<DrawerBoxPart> parts = new List<DrawerBoxPart>();

			MaterialType frontbackMat = SideMaterial == MaterialType.HybridBirch ? MaterialType.EconomyBirch : SideMaterial;
			MaterialType sideMat = SideMaterial == MaterialType.HybridBirch ? MaterialType.SolidBirch : SideMaterial;

			DrawerBoxPart front = new DrawerBoxPart {
				CutListName = "Front",
				PartType = DBPartType.Side,
				UseType = InventoryUseType.Linear,
				Material = frontbackMat,
				Qty = Qty,
				Width = Height,
				Length = Width + ManufacturingConstants.FrontBackAdj
			};
			parts.Add(front);

			DrawerBoxPart backLeft = new DrawerBoxPart {
				CutListName = "Back Left - A",
				PartType = DBPartType.Side,
				UseType = InventoryUseType.Linear,
				Material = frontbackMat,
				Qty = Qty,
				Width = Height,
				Length = A + ManufacturingConstants.FrontBackAdj
			};
			parts.Add(backLeft);

			DrawerBoxPart backCenter = new DrawerBoxPart {
				CutListName = "Back Center",
				PartType = DBPartType.Side,
				UseType = InventoryUseType.Linear,
				Material = frontbackMat,
				Qty = Qty,
				Width = Height,
				Length = Width - A - B + 2 * ManufacturingConstants.SideThickness + ManufacturingConstants.FrontBackAdj
			};
			parts.Add(backCenter);

			DrawerBoxPart backRight = new DrawerBoxPart {
				CutListName = "Back Right - B",
				PartType = DBPartType.Side,
				UseType = InventoryUseType.Linear,
				Material = frontbackMat,
				Qty = Qty,
				Width = Height,
				Length = B + ManufacturingConstants.FrontBackAdj
			};
			parts.Add(backRight);

			DrawerBoxPart sides = new DrawerBoxPart {
				CutListName = "Sides",
				PartType = DBPartType.Side,
				UseType = InventoryUseType.Linear,
				Material = sideMat,
				Qty = Qty * 2,
				Width = Height,
				Length = Depth - ManufacturingConstants.SideThickness
			};
			parts.Add(sides);

			DrawerBoxPart sidesCenter = new DrawerBoxPart {
				CutListName = "Sides Center - C",
				PartType = DBPartType.Side,
				UseType = InventoryUseType.Linear,
				Material = sideMat,
				Qty = Qty,
				Width = Height,
				Length = C
			};
			parts.Add(sidesCenter);

			DrawerBoxPart bottom = new DrawerBoxPart {
				CutListName = "Bottom",
				PartType = DBPartType.Bottom,
				UseType = InventoryUseType.Area,
				Material = BottomMaterial,
				Qty = Qty,
				Width = Width - 2 * ManufacturingConstants.SideThickness + 2 * ManufacturingConstants.DadoDepth - ManufacturingConstants.BottomAdj,
				Length = Depth - 2 * ManufacturingConstants.SideThickness + 2 * ManufacturingConstants.DadoDepth - ManufacturingConstants.BottomAdj
			};
			parts.Add(bottom);


			return parts;

		}

	}
}
