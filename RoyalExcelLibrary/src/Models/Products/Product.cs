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

		public int LineNumber { get; set; }

		public abstract IEnumerable<Part> GetParts();

	}
}
