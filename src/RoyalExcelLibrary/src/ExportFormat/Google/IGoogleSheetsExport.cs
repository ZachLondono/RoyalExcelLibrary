using RoyalExcelLibrary.ExcelUI.Models;

namespace RoyalExcelLibrary.ExcelUI.ExportFormat.Google {

	public interface IGoogleSheetsExport {
		void ExportOrder(Order order);

	}

}
