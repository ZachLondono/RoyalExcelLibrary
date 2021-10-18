using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Models {
	
	// <summary>
	// A single part which makes up a composite part or a product
	// </summary>
	public class Part {
		public MaterialType Material { get; set; }
		public InventoryUseType UseType { get; set; }
		public double Width { get; set; }
		public double Length { get; set; }
		public int Qty { get; set; }
	}
}
