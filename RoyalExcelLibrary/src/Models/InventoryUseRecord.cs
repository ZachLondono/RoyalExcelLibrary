using RoyalExcelLibrary.DAL.Repositories;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Models {
	public class InventoryUseRecord  : BaseRepoClass {

		public int ItemId { get; set; }
		
		public int JobId { get; set; }

		public int Qty { get; set; }

		public DateTime Timestamp { get; set; }

	}

}
