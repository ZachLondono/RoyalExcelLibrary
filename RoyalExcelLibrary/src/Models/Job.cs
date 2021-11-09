using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RoyalExcelLibrary.DAL.Repositories;
using RoyalExcelLibrary.Models.Products;

namespace RoyalExcelLibrary.Models {
	public class Job : BaseRepoClass {

		public string JobSource { get; set; }

		public Status Status { get; set; }

		public string Name { get; set; }

		public DateTime CreationDate { get; set; }

		public decimal GrossRevenue { get; set; }

	}
}
