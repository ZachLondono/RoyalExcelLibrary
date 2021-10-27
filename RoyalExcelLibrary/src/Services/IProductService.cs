using RoyalExcelLibrary.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Services {
	interface IProductService {

		// <summary>
		//  Stores the order and its items in there respective repositories and returns the same order with it's ID set 
		// </summary>
		Order StoreCurrentOrder(Order order);

		Microsoft.Office.Interop.Excel.Worksheet[] GenerateCutList(Order order, Microsoft.Office.Interop.Excel.Workbook outputSheet);

		void GenerateConfirmation();

		void ConfirmOrder();

		void GenerateInvoice();

		void PayOrder();

	}
}
