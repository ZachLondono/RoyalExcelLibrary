using RoyalExcelLibrary.Models;

namespace RoyalExcelLibrary.ExportFormat.Google {

	public interface IGoogleSheetsExport {
		void ExportOrder(Order order);

	}

}
