using RoyalExcelLibrary.DAL.Repositories;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Models.Products {
	public class DrawerBox : BaseRepoClass, IProduct {

		public double Height { get; set; }
		public double Width { get; set; }
		public double Depth { get; set; }
		public int Qty { get; set; }
		public MaterialType SideMaterial { get; set; }
		public MaterialType BottomMaterial { get; set; }
		public int JobId { get; set; }

		public IEnumerable<Part> GetParts() {

			double DadoDepth = 1;
			double DovetailDepth = 2;

			List<Part> parts = new List<Part>();

			Part front = new Part();
			front.Qty = Qty * 2;
			front.Width = Height;
			front.Length = Width;
			front.UseType = InventoryUseType.Linear;
			if (SideMaterial == MaterialType.HybridBirch)
				front.Material = MaterialType.EconomyBirch;
			else front.Material = SideMaterial;

			Debug.WriteLine($"Front: {front.Width}H x {front.Length}L");

			Part sides = new Part();
			sides.Qty = Qty * 2;
			sides.Width = Height;
			sides.Length = Depth - 2 * DovetailDepth;
			sides.UseType = InventoryUseType.Linear;
			if (SideMaterial == MaterialType.HybridBirch)
				sides.Material = MaterialType.SolidBirch;
			else sides.Material = SideMaterial;

			Debug.WriteLine($"Sides: {sides.Width}H x {sides.Length}L");

			Part bottom = new Part();
			bottom.Width = Width - 2 * DadoDepth;
			bottom.Length = Depth - 2 * DadoDepth;
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
