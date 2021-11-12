using RoyalExcelLibrary.DAL.Repositories;
using RoyalExcelLibrary.ExportFormat;
using RoyalExcelLibrary.Models.Products;
using RoyalExcelLibrary.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.Models {
	public class Order : BaseRepoClass {

		public Job Job { get; private set; }

		public Status Status { get; set; }

		public string Number { get; set; }

		public decimal SubTotal { get; set; }

		public decimal Tax { get; set; }

		public decimal ShippingCost { get; set; }

		public IList<string> InfoFields { get; set; }

		public Company Customer { get; set; }

		public Company Vendor { get; set; }

		public Company Supplier { get; set; }


		public IEnumerable<Product> Products {
			get { return _products; }
		}

		private readonly List<Product> _products;

		public Order(Job job) {
			Job = job;
			_products = new List<Product>();
		}

		public void AddProduct(Product product) {
			_products.Add(product);
		}

		public void AddProducts(IEnumerable<Product> products) {
			_products.AddRange(products);
		}

	}

    public class HafeleOrder : Order {
		public string ProjectNumber { get; set; }
		public string ProNumber { get; set; }
		public string ConfigNumber { get; set; }
		public string ClientAccountNumber { get; set; }
		public string ClientPurchaseOrder { get; set; }
        public HafeleOrder(Job job) : base(job) {

			Supplier = new Company();
			Supplier.Name = "Royal Cabinet Co.";
			Supplier.Address = new Address {
				Line1 = "15E Easy St",
				City = "Bound Brook",
				State = "NJ",
				Zip = "08805"
			};

			Vendor = new Company();
			Vendor.Name = "Hafele America Co.";
			Vendor.Address = new Address {
				Line1 = "3901 Cheyenne Drive",
				City = "Archdale",
				State = "NC",
				Zip = "27263",
			};

		}
    }

    public class RichelieuOrder : Order {
		public string ClientFirstName { get; set; }
		public string ClientLastName { get; set; }
		public string RichelieuNumber { get; set; }
		public string ClientPurchaseOrder { get; set; }
		public string WebNumber { get; set; }
        public RichelieuOrder(Job job) : base(job) {

			Supplier = new Company();
			Supplier.Name = "Royal Cabinet Co.";
			Supplier.Address = new Address {
				Line1 = "15E Easy St",
				City = "Bound Brook",
				State = "NJ",
				Zip = "08805"
			};

			Vendor = new Company();
			Vendor.Name = "Richelieu America ltd";
			Vendor.Address = new Address {
				Line1 = "",
				City = "",
				State = "",
				Zip = "",
			};

		}
    }


}
