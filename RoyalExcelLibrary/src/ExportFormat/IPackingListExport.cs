using Microsoft.Office.Interop.Excel;
using RoyalExcelLibrary.Models;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace RoyalExcelLibrary.ExportFormat {

	public class PackingListData {

		public string SupplierName { get; set; }
		public Address SupplierAddress { get; set; }
		public string SupplierContact { get; set; }

		public string RecipientName { get; set; }
		public Address RecipientAddress { get; set; }
		public string RecipientContact { get; set; }

	}

	public class Address {
		public string StreetAddress { get; set; }
		public string City { get; set; }
		public string State { get; set; }
		public string Zip { get; set; }
	}

	public interface IPackingListExport {

		Worksheet ExportOrder(Order order, PackingListData data, Workbook workbook);

	}

}
