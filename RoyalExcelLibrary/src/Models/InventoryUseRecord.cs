using System;

namespace RoyalExcelLibrary.Models {
	public class InventoryUseRecord {
		
		public int Id { get; set; }
		public int JobId { get; set; }

		public int Qty { get; set; }

		public MaterialType Material { get; set; }

		public double Width { get; set; }

		public double Length { get; set; }

		public double Thickness { get; set; }

		public DateTime Timestamp { get; set; }

	}

}
