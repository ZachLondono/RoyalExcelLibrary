using RoyalExcelLibrary.ExcelUI.Models;
using RoyalExcelLibrary.ExcelUI.Services;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.ExcelUI.ExportFormat.Labels {

	public interface ILabelExport {

		void PrintLables(Order order, ILabelServiceFactory factory);

	}

}
