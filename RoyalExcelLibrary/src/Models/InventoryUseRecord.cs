using RoyalExcelLibrary.DAL.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Models {
	public class InventoryUseRecord  : BaseRepoClass {
		
		public int JobId { get; set; }

		public int Qty { get; set; }

		public MaterialType Material { get; set; }

		public double Width { get; set; }

		public double Length { get; set; }

		public double Thickness { get; set; }

		public DateTime Timestamp { get; set; }

	}

}
