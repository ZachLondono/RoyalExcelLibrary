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

	public class DrawerBox : BaseRepoClass, IProduct {
		public int JobId { get; set; }
		public double Height { get; set; }
		public double Width { get; set; }
		public double Depth { get; set; }
		public int Qty { get; set; }
		public MaterialType SideMaterial { get; set; }
		public MaterialType BottomMaterial { get; set; }
		public Clips ClipsOption { get; set; }
		public UndermountNotch NotchOption { get; set; }
		public Insert InsertOption { get; set; }
		public bool MountingHoles { get; set; }
		public bool ScoopFront { get; set; }
		public bool Logo { get; set; }
		public bool PostFinish { get; set; }
		public string LabelNote { get; set; }

		public virtual IEnumerable<Part> GetParts() {

			List<Part> parts = new List<Part>();

			DrawerBoxPart front = new DrawerBoxPart();
			front.PartType = DBPartType.Side;
			front.CutListName = "Front/Back";
			front.Qty = Qty * 2;
			front.Width = Height;
			front.Length = Width;
			front.UseType = InventoryUseType.Linear;
			if (SideMaterial == MaterialType.HybridBirch)
				front.Material = MaterialType.EconomyBirch;
			else front.Material = SideMaterial;

			DrawerBoxPart sides = new DrawerBoxPart();
			sides.PartType = DBPartType.Side;
			sides.CutListName = "Sides";
			sides.Qty = Qty * 2;
			sides.Width = Height;
			sides.Length = Depth - 2 * ManufacturingConstants.DovetailDepth;
			sides.UseType = InventoryUseType.Linear;
			if (SideMaterial == MaterialType.HybridBirch)
				sides.Material = MaterialType.SolidBirch;
			else sides.Material = SideMaterial;

			DrawerBoxPart bottom = new DrawerBoxPart();
			bottom.PartType = DBPartType.Bottom;
			bottom.CutListName = "Bottom";
			bottom.Width = Width - 2 * ManufacturingConstants.DadoDepth;
			bottom.Length = Depth - 2 * ManufacturingConstants.DadoDepth;
			bottom.Qty = Qty;
			bottom.UseType = InventoryUseType.Area;
			bottom.Material = BottomMaterial;

			parts.Add(front);
			parts.Add(sides);
			parts.Add(bottom);

			return parts;

		}

	}

}
