using RoyalExcelLibrary.DAL.Repositories;
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
		public Insert InsertOption { get; set; }
		public bool MountingHoles { get; set; }
		public bool ScoopFront { get; set; }
		public bool Logo { get; set; }
		public bool PostFinish { get; set; }

		public override IEnumerable<Part> GetParts() {

			List<Part> parts = new List<Part>();

			DrawerBoxPart front = new DrawerBoxPart();
			front.PartType = DBPartType.Side;
			front.Qty = Qty * 2;
			front.Width = Height;
			front.Length = Width + ManufacturingConstants.FrontBackAdj;
			front.UseType = InventoryUseType.Linear;

			if ((SideMaterial == MaterialType.HybridBirch || SideMaterial == MaterialType.EconomyBirch) && ScoopFront) {
				DrawerBoxPart back = new DrawerBoxPart();
				back.PartType = DBPartType.Side;
				back.Qty = Qty * 2;
				back.Width = Height;
				back.Length = Width + ManufacturingConstants.FrontBackAdj;
				back.UseType = InventoryUseType.Linear;
				back.CutListName = "Back";
				back.Material = MaterialType.EconomyBirch;

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

			DrawerBoxPart sides = new DrawerBoxPart();
			sides.PartType = DBPartType.Side;
			sides.CutListName = "Sides";
			sides.Qty = Qty * 2;
			sides.Width = Height;
			sides.Length = Depth - ManufacturingConstants.SideAdj;
			sides.UseType = InventoryUseType.Linear;
			if (SideMaterial == MaterialType.HybridBirch)
				sides.Material = MaterialType.SolidBirch;
			else sides.Material = SideMaterial;

			parts.Add(sides);

			DrawerBoxPart bottom = new DrawerBoxPart();
			bottom.PartType = DBPartType.Bottom;
			bottom.CutListName = "Bottom";
			bottom.Width = Width - 2 * ManufacturingConstants.SideThickness + 2 * ManufacturingConstants.DadoDepth - ManufacturingConstants.BottomAdj;
			bottom.Length = Depth - 2 * ManufacturingConstants.SideThickness + 2 * ManufacturingConstants.DadoDepth - ManufacturingConstants.BottomAdj;
			bottom.Qty = Qty;
			bottom.UseType = InventoryUseType.Area;
			bottom.Material = BottomMaterial;

			parts.Add(bottom);

			return parts;

		}

	}

}
