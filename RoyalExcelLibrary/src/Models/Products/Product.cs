using RoyalExcelLibrary.DAL.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Models.Products {
	public abstract class Product : BaseRepoClass {

		public int Qty { get; set; }

		public double UnitPrice { get; set; }

		// THe line number of the item in the customer's order 
		public int LineNumber { get; set; }

		public IList<string> InfoFields { get; set; }

		// Returns a list of all the parts needed for the item
		public abstract IEnumerable<Part> GetParts();

	}
}
