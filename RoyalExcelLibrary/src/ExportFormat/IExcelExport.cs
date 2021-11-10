using Microsoft.Office.Interop.Excel;
using RoyalExcelLibrary.Models;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace RoyalExcelLibrary.ExportFormat {

	public class ExportData {

		public string SupplierName { get; set; }
		public Address SupplierAddress { get; set; }
		public string SupplierContact { get; set; }

		public string RecipientName { get; set; }
		public Address RecipientAddress { get; set; }
		public string RecipientContact { get; set; }

	}

	public class Address {
		public string Line1 { get; set; }
		public string Line2 { get; set; }
		public string City { get; set; }
		public string State { get; set; }
		public string Zip { get; set; }

		public override string ToString() {
			if (string.IsNullOrEmpty(Line2))
				return $"{Line1}\n{City}, {State} {Zip}";
			return $"{Line1}\n{Line2}\n{City}, {State} {Zip}";
		}

	}

	public interface IExcelExport {

		Worksheet ExportOrder(Order order, Workbook workbook);

	}

}
