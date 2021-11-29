using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Views;

using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace RoyalExcelLibrary.Services {
	interface IProductService {

		Dictionary<string,Worksheet> GenerateCutList(Order order, Workbook outputBook, ErrorMessage errorOutput);

		Worksheet GenerateConfirmation(Order order, Workbook outputBook, ErrorMessage errorOutput);

		Worksheet GenerateInvoice(Order order, Workbook outputBook, ErrorMessage errorOutput);

		Worksheet GeneratePackingList(Order order, Workbook outputBook, ErrorMessage errorOutput);

	}
}
