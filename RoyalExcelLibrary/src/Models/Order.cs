using RoyalExcelLibrary.ExportFormat;
using RoyalExcelLibrary.Models.Products;
using System.Collections.Generic;

namespace RoyalExcelLibrary.Models {
	public class Order {

		public int Id { get; set; }

		public Job Job { get; private set; }

		public string Number { get; set; }

		public decimal SubTotal { get; set; }

		public decimal Tax { get; set; }

		public decimal ShippingCost { get; set; }

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
		public string SourceFile { get; set; }
        public HafeleOrder(Job job) : base(job) {

            Supplier = new Company {
                Name = "Royal Cabinet Co.",
                Address = new Address {
                    Line1 = "15E Easy St",
                    City = "Bound Brook",
                    State = "NJ",
                    Zip = "08805"
                }
            };

            Vendor = new Company {
                Name = "Hafele America Co.",
                Address = new Address {
                    Line1 = "3901 Cheyenne Drive",
                    City = "Archdale",
                    State = "NC",
                    Zip = "27263",
                }
            };

        }
    }

    public class RichelieuOrder : Order {
		public string ClientFirstName { get; set; }
		public string ClientLastName { get; set; }
		public string RichelieuNumber { get; set; }
		public string ClientPurchaseOrder { get; set; }
		public string WebNumber { get; set; }
        public string CustomerNum { get; set; }

        public RichelieuOrder(Job job) : base(job) {

            Supplier = new Company {
                Name = "Royal Cabinet Co.",
                Address = new Address {
                    Line1 = "15E Easy St",
                    City = "Bound Brook",
                    State = "NJ",
                    Zip = "08805"
                }
            };

            Vendor = new Company {
                Name = "Richelieu America ltd",
                Address = new Address {
                    Line1 = "",
                    City = "",
                    State = "",
                    Zip = "",
                }
            };

        }
    }


}
