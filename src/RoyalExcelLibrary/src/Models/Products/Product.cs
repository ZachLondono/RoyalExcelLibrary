using System.Collections.Generic;

namespace RoyalExcelLibrary.ExcelUI.Models.Products {
	public abstract class Product {

		public int Id { get; set; }

		public int Qty { get; set; }

		public string ProductName { get; set; }

		public string ProductDescription { get; set; }

		public decimal UnitPrice { get; set; }

		// THe line number of the item in the customer's order 
		public int LineNumber { get; set; }

		public string LevelName { get; set; }

		public string Note { get; set; }

		// Returns a list of all the parts needed for the item
		public abstract IEnumerable<Part> GetParts();

	}
}
