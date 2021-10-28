using RoyalExcelLibrary.DAL.Repositories;
using RoyalExcelLibrary.ExportFormat;
using RoyalExcelLibrary.Models.Products;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Models {
	public class Order : BaseRepoClass {

		public Job Job { get; private set; }

		public Status Status { get; set; }

		public string CustomerName { get; set; }

		public string Number { get; set; }

		public Address ShipAddress { get; set; }

		public IEnumerable<IProduct> Products {
			get { return _products; }
		}

		private readonly List<IProduct> _products;

		public Order(Job job, string customerName, string number) {
			Job = job;
			CustomerName = customerName;
			Number = number;
			Status = Status.UnConfirmed;
			_products = new List<IProduct>();
		}

		public void AddProduct(IProduct product) {
			_products.Add(product);
		}

		public void AddProducts(IEnumerable<IProduct> products) {
			_products.AddRange(products);
		}

	}

}
