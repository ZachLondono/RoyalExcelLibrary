using RoyalExcelLibrary.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Services {
	interface IProductService {

		void StoreCurrentOrder(Order order);

		void GenerateConfirmation();

		void ConfirmOrder();

		void GenerateInvoice();

		void PayOrder();

	}
}
