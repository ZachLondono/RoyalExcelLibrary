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

		public double SubTotal { get; set; }
		
		public double Tax { get; set; }

		public double ShippingCost { get; set; }

		// Extra meta-info relating to the order
		public IEnumerable<string> InfoFields { get; set; }

		public IEnumerable<Product> Products {
			get { return _products; }
		}

		private readonly List<Product> _products;

		public Order(Job job, string customerName, string number) {
			Job = job;
			CustomerName = customerName;
			Number = number;
			Status = Status.UnConfirmed;
			_products = new List<Product>();
		}

		public void AddProduct(Product product) {
			_products.Add(product);
		}

		public void AddProducts(IEnumerable<Product> products) {
			_products.AddRange(products);
		}

	}

}
