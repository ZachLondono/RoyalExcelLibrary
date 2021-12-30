using RoyalExcelLibrary.ExcelUI.ExportFormat;
using RoyalExcelLibrary.ExcelUI.Models.Products;
using System.Collections.Generic;

namespace RoyalExcelLibrary.ExcelUI.Models {
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

        public bool Rush { get; set; } = false;

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
                    Line2 = "P.O. Box 4000",
                    City = "Archdale",
                    State = "NC",
                    Zip = "27263",
                }
            };

        }
    }

    public class AllmoxyOrder : Order {

        public string OrderDescription { get; set; }
        public string OrderNote { get; set; }

        public string ShippingInstructions { get; set; }

        public AllmoxyOrder(Job job) : base(job) {

            Address royalAddress = new Address {
                Line1 = "15E Easy St",
                Line2 = "",
                City = "Bound Brook",
                State = "NJ",
                Zip = "08805"
            };

            Vendor = new Company {
                Name = "OT",
                Address = royalAddress
            };

            Supplier = new Company {
                Name = "Metro Cabinet Parts",
                Address = royalAddress
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
                Name = "Richelieu America Ltd.",
                Address = new Address {
                    Line1 = "7021 Sterling Ponds Blvd.",
                    City = "Sterling Heights",
                    State = "MI",
                    Zip = "48312",
                }
            };

        }
    }


}
