using RoyalExcelLibrary.Models.Products;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Models {
	public class Order {
	
		public Job Job { get; private set; }
		public IEnumerable<IProduct> Products {
			get { return _products; }
		}

		private readonly List<IProduct> _products;

		public Order(Job job) {
			Job = job;
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
