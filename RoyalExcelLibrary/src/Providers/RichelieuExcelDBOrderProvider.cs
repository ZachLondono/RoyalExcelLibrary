using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System.Diagnostics;
using System.Xml;
using RoyalExcelLibrary.Models.Options;
using System.Security.Cryptography.X509Certificates;
using System.Net;
using System.IO;

namespace RoyalExcelLibrary.Providers {
	class RichelieuExcelDBOrderProvider : IOrderProvider {

		private readonly string _webnumber;

		public RichelieuExcelDBOrderProvider(string webnumber) {
			_webnumber = webnumber;
		}

		public Order LoadCurrentOrder() {


			string content;
			try {
				HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create("https://xml.richelieu.com/royalCabinet/getOrderDetails.php?id=" + _webnumber);

				// Find richelieu certificate by thumbprint, stored in LocalMachine\My
				X509Store store = new X509Store(StoreName.My, StoreLocation.LocalMachine);
				store.Open(OpenFlags.ReadOnly);

				X509Certificate rich_cert = null;
				foreach (var cert in store.Certificates) {
					if (cert.Thumbprint.ToUpper().Equals("35430E729F268ACE03C7B3FA3F443F0822C5F9F7")) {
						rich_cert = cert;
						break;
					}
				}

				if (rich_cert is null) throw new InvalidOperationException("Richelieu api certificate unavailable");
				
				request.ClientCertificates.Add(rich_cert);
				HttpWebResponse response = (HttpWebResponse)request.GetResponse();

				using (var reader = new StreamReader(response.GetResponseStream())) {
					content = reader.ReadToEnd();
				}
			} catch (Exception e) {
				throw new InvalidOperationException("Can't Download Order from Richelieu", e);
			}

			
			XmlDocument doc = new XmlDocument();
			doc.LoadXml(content);

			var _currentOrderNode = doc.FirstChild;
			if (_currentOrderNode.LocalName.Equals("xml")) {
				_currentOrderNode = _currentOrderNode.NextSibling;
			}
			_currentOrderNode = _currentOrderNode.FirstChild;

			XmlNode shippingNode = _currentOrderNode["shipTo"];
			XmlAttributeCollection attributes = shippingNode.Attributes;
			string company = attributes.GetNamedItem("company").InnerText;
			string streetAddress = attributes.GetNamedItem("address1").InnerText;
			string streetAddress2 = attributes.GetNamedItem("address2").InnerText;
			string city = attributes.GetNamedItem("city").InnerText;
			string state = attributes.GetNamedItem("province").InnerText;
			string zip = attributes.GetNamedItem("postalCode").InnerText;
			string firstName = attributes.GetNamedItem("firstName").InnerText;
			string lastName = attributes.GetNamedItem("lastName").InnerText;
			string customerNum = attributes.GetNamedItem("richelieuNumber").InnerText;

			XmlNode headerNode = _currentOrderNode["header"];
			attributes = headerNode.Attributes;
			string creationDate = attributes.GetNamedItem("orderDate").InnerText;
			string webOrder = attributes.GetNamedItem("webOrder").InnerText;
			string richelieuOrder = attributes.GetNamedItem("richelieuOrder").InnerText;
			string richelieuPO = attributes.GetNamedItem("richelieuPO").InnerText;
			string clientPO = attributes.GetNamedItem("clientPO").InnerText;

			Job job = new Job {
				CreationDate = DateTime.Parse(creationDate),
				GrossRevenue = 0,
				JobSource = "Richelieu",
				Name = clientPO
			};

            RichelieuOrder order = new RichelieuOrder(job) {
                ShippingCost = 0,
                Tax = 0,
                Number = richelieuPO,
                ClientFirstName = firstName,
                ClientLastName = lastName,
                RichelieuNumber = richelieuOrder,
                WebNumber = webOrder,
				CustomerNum = customerNum,
				Customer = new Company {
                    Name = company,
                    Address = new ExportFormat.Address {
                        Line1 = streetAddress,
                        Line2 = streetAddress2,
                        City = city,
                        State = state,
                        Zip = zip
                    }
                }
            };

            var linesNodes = _currentOrderNode.SelectNodes("/response/order/line");
			int line = 0;
			foreach (XmlNode linesNode in linesNodes) {
				string description = linesNode.Attributes.GetNamedItem("descriptionEn").InnerText;
				string sku = linesNode.Attributes.GetNamedItem("sku").InnerText;

				string[] properties = description.Split(',');
				MaterialType sideMat = ParseMaterial(properties[1].Trim());
				MaterialType bottMat = ParseMaterial(properties[3].Trim());
				UndermountNotch notch = ParseNotch(properties[5].Trim());
				Clips clips = Clips.No_Clips;
				bool scoopFront = !properties[8].Trim().Equals("Standard Drawer - No Pull-Out");

				string note = linesNode.Attributes.GetNamedItem("note").InnerText;
				if (!string.IsNullOrWhiteSpace(note))
					System.Windows.Forms.MessageBox.Show(note, "Order Note");

				XmlNodeList boxNodes = linesNode.SelectNodes($"/response/order/line[{++line}]/dimension");

				int lineNum = 1;
				foreach (XmlNode dimension in boxNodes) {

					string qty_str = dimension.Attributes.GetNamedItem("qty").InnerText;
					string height_str = dimension.Attributes.GetNamedItem("HEIGHT").InnerText;  // Comes in mm
					string width_str = dimension.Attributes.GetNamedItem("WIDTH").InnerText;    // Comes in inches
					string depth_str = dimension.Attributes.GetNamedItem("DEPTH").InnerText;    // Comes in inches
					string unitPrice_str = dimension.Attributes.GetNamedItem("price").InnerText;

                    DrawerBox box = new DrawerBox {
                        SideMaterial = sideMat,
                        BottomMaterial = bottMat,
                        Qty = Convert.ToInt32(qty_str),
                        Height = Convert.ToDouble(height_str),
                        Width = FractionToDouble(width_str) * 25.4,
                        Depth = FractionToDouble(depth_str) * 25.4,
                        UnitPrice = Convert.ToDecimal(unitPrice_str),
                        ClipsOption = clips,
                        NotchOption = notch,
                        MountingHoles = false,
                        InsertOption = "",
                        Logo = false,
                        PostFinish = false,
                        ScoopFront = scoopFront,
                        LineNumber = lineNum++,

                        Note = note,
                        ProductName = sku,
                        ProductDescription = $"{properties[1]}\n{properties[3]}\n{properties[5]}\n{properties[3]}\n{properties[6]}\n{properties[8]}"
                    };

                    order.AddProduct(box);

				}
			}

			order.SubTotal = order.Products.Sum(b => Convert.ToDecimal(b.Qty) * b.UnitPrice);

			return order;
		}

		private double FractionToDouble(string fraction) {

			string[] parts = fraction.Split(' ', '/');

			double val = Convert.ToDouble(parts[0]);
			if (parts.Length == 3) {

				double numerator = Convert.ToDouble(parts[1]);
				double denomenator = Convert.ToDouble(parts[2]);

				val += numerator / denomenator;

			}

			return val;

		}

		private UndermountNotch ParseNotch(string name) {

			switch(name) {

				case "Standard Back Notch with Drilling for Hook":
					return UndermountNotch.Std_Notch;
				case "Front (96 mm) and back (37 mm) notch":
					return UndermountNotch.Front_Back;
				case "No Notch":
					return UndermountNotch.No_Notch;
				default:
					return UndermountNotch.Unknown;

			}

		}

		private MaterialType ParseMaterial(string name) {

			switch (name) {
				case "Economy Birch (Finger Jointed)":
					return MaterialType.EconomyBirch;
				case "Solid Birch (No Finger Joint)":
					return MaterialType.SolidBirch;
				case "Solid Birch (No Finger Joint) - SIDES ONLY":
					return MaterialType.HybridBirch;
				case "Walnut":
					return MaterialType.SolidWalnut;
				case "1/4\" Bottom":
					return MaterialType.Plywood1_4;
				case "1/2\" Bottom":
					return MaterialType.Plywood1_2;
				default:
					return MaterialType.Unknown;
			}

		}

	}

}