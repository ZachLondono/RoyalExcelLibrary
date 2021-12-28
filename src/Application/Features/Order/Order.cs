using RoyalExcelLibrary.Application.Features.Product;
using System;
using System.Collections.Generic;

namespace RoyalExcelLibrary.Application.Features.Order {
    public class Order {

        public int Id { get; set; }

        /// <summary>
        /// The name of the company placing the order
        /// </summary>
        public string Customer { get; set; }
        
        /// <summary>
        /// The name of the order
        /// </summary>
        public string OrderName { get; set; }

        /// <summary>
        /// The reference number of the order
        /// </summary>
        public string OrderNumber { get; set; }

        /// <summary>
        /// The invoice information for the order
        /// </summary>
        public Invoice Invoice { get; set; }

        public DateTime OrderDate { get; set; }

        /// <summary>
        /// Extra order details that can change depending on the order source
        /// </summary>
        public Dictionary<string, string> OrderDetails { get; set; } = new Dictionary<string, string>();

        /// <summary>
        /// The products in the order
        /// </summary>
        public IEnumerable<IProduct> Products { get; set; } 

    }

    public class Invoice {
        public Address BillingAddress { get; private set; }
        public decimal SubTotal { get; private set; }
        public decimal Tax { get; private set; }
        public decimal Shipping { get; private set; }
        
        public Invoice(Address billingAddress, decimal subTotal, decimal tax, decimal shipping) {
            BillingAddress = billingAddress;
            SubTotal = subTotal;
            Tax = tax;
            Shipping = shipping;
        }
    }

    public class Address {
        public string Line1 { get; private set; }
        public string Line2 { get; private set; }
        public string City { get; private set; }
        public string State { get; private set; }
        public string Zip { get; private set; }

        public override bool Equals(object obj) {
            if (obj == null || !(obj is Address)) return false;
            return ToString().ToLower().Equals((obj as Address).ToString().ToLower());
        }

        public override int GetHashCode() {
            return Line1.GetHashCode() | Line2.GetHashCode() | City.GetHashCode() | State.GetHashCode() | Zip.GetHashCode();
        }

        public override string ToString() {
            return $"{Line1}, {Line2}, {City}, {State}, {Zip}";
        }
    }

}
