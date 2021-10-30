using RoyalExcelLibrary.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;

namespace RoyalExcelLibrary.Services {
	interface IProductService {

		// <summary>
		//  Stores the order and its items in there respective repositories and returns the same order with it's ID set 
		// </summary>
		Order StoreCurrentOrder(Order order);

		void SetOrderStatus(Order order, Status status);

		Worksheet[] GenerateCutList(Order order, Workbook outputBook);

		Worksheet GenerateConfirmation(Order order, Workbook outputBook);

		Worksheet GenerateInvoice(Order order, Workbook outputBook);

		Worksheet GeneratePackingList(Order order, Workbook outputBook);

	}
}
