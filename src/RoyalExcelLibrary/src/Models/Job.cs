using System;

namespace RoyalExcelLibrary.ExcelUI.Models {
	public class Job {

		public int Id { get; set; }

		public string JobSource { get; set; }

		public string Name { get; set; }

		public DateTime CreationDate { get; set; }

		public decimal GrossRevenue { get; set; }

	}
}
