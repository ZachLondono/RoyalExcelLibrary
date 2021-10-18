using RoyalExcelLibrary.DAL.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Models.Products {
	public interface IProduct {
		IEnumerable<Part> GetParts();
	}
}
