using RoyalExcelLibrary.ExcelUI.Models;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace RoyalExcelLibrary.ExcelUI.ExportFormat {
	public interface ICutListFormat {

        Excel.Range WriteOrderHeader(Order order, Excel.Worksheet outputsheet);

        Excel.Range WriteOrderParts(IEnumerable<string[,]> seperatedBoxes, Excel.Worksheet outputsheet, int startRow, int startCol);

    }

}
