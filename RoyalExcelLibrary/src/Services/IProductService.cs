using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Views;

using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace RoyalExcelLibrary.Services {
	interface IProductService {

		// <summary>
		//  Stores the order and its items in there respective repositories and returns the same order with it's ID set 
		// </summary>
		Order StoreCurrentOrder(Order order);

		void SetOrderStatus(Order order, Status status);

		Dictionary<string,Worksheet> GenerateCutList(Order order, Workbook outputBook, ErrorMessage errorOutput);

		Worksheet GenerateConfirmation(Order order, Workbook outputBook, ErrorMessage errorOutput);

		Worksheet GenerateInvoice(Order order, Workbook outputBook, ErrorMessage errorOutput);

		Worksheet GeneratePackingList(Order order, Workbook outputBook, ErrorMessage errorOutput);

	}
}
