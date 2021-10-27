using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Models.Products {
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
				Length = Width
			};
			parts.Add(front);

			DrawerBoxPart backLeft = new DrawerBoxPart {
				CutListName = "Back Left - A",
				PartType = DBPartType.Side,
				UseType = InventoryUseType.Linear,
				Material = frontbackMat,
				Qty = Qty,
				Width = Height,
				Length = A
			};
			parts.Add(backLeft);

			DrawerBoxPart backCenter = new DrawerBoxPart {
				CutListName = "Back Center",
				PartType = DBPartType.Side,
				UseType = InventoryUseType.Linear,
				Material = frontbackMat,
				Qty = Qty,
				Width = Height,
				Length = Width - A - B + 2 * ManufacturingConstants.DovetailDepth + 1
			};
			parts.Add(backCenter);

			DrawerBoxPart backRight = new DrawerBoxPart {
				CutListName = "Back Right - B",
				PartType = DBPartType.Side,
				UseType = InventoryUseType.Linear,
				Material = frontbackMat,
				Qty = Qty,
				Width = Height,
				Length = B
			};
			parts.Add(backRight);

			DrawerBoxPart sides = new DrawerBoxPart {
				CutListName = "Sides",
				PartType = DBPartType.Side,
				UseType = InventoryUseType.Linear,
				Material = sideMat,
				Qty = Qty * 2,
				Width = Height,
				Length = Depth - 2 * ManufacturingConstants.DovetailDepth
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
				Width = Width - 2 * ManufacturingConstants.DadoDepth,
				Length = Depth - 2 * ManufacturingConstants.DadoDepth
			};
			parts.Add(bottom);


			return parts;

		}

	}
}
