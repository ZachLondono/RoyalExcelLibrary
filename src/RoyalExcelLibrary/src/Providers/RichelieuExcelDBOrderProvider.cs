using System;
using System.Linq;

using RoyalExcelLibrary.ExcelUI.Models;
using RoyalExcelLibrary.ExcelUI.Models.Products;
using System.Xml;
using RoyalExcelLibrary.ExcelUI.Models.Options;
using System.Security.Cryptography.X509Certificates;
using System.Net;
using System.IO;

namespace RoyalExcelLibrary.ExcelUI.Providers {
	public class RichelieuExcelDBOrderProvider : IOrderProvider {

		public string XMLContent { get; set; }

		public void DownloadOrder(string webnumber) {

			try {
				HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create("https://xml.richelieu.com/royalCabinet/getOrderDetails.php?id=" + webnumber);

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
					XMLContent = reader.ReadToEnd();
				}
			} catch (Exception e) {
				throw new InvalidOperationException("Can't Download Order from Richelieu", e);
			}

		}

		public Order LoadCurrentOrder() {
			
			if (string.IsNullOrEmpty(XMLContent)) {
				throw new InvalidOperationException("No order data loaded");
            }

			XmlDocument doc = new XmlDocument();
			doc.LoadXml(XMLContent);

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
                Number = richelieuOrder,
                ClientFirstName = firstName,
                ClientLastName = lastName,
                RichelieuNumber = richelieuPO,
                WebNumber = webOrder,
				CustomerNum = customerNum,
				ClientPurchaseOrder = clientPO,
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
				string[] properties = description.Split(',');

				string sku = linesNode.Attributes.GetNamedItem("sku").InnerText;

				RichelieuConfiguration config = ParseSku(sku);
				Clips clips = Clips.No_Clips;

				order.Rush = config.Rush;

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
                        Qty = Convert.ToInt32(qty_str),
                        Height = Convert.ToDouble(height_str),
                        Width = FractionToDouble(width_str) * 25.4,
                        Depth = FractionToDouble(depth_str) * 25.4,
                        UnitPrice = Convert.ToDecimal(unitPrice_str),
                        ClipsOption = clips,
						SideMaterial = config.BoxMaterial,
						BottomMaterial = config.BotMaterial,
						NotchOption = config.Notch,
                        MountingHoles = false,
                        InsertOption = "",
                        Logo = false,
                        PostFinish = false,
                        ScoopFront = config.ScoopFront,
						PullOutFront = config.PullOutFront,
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
				case "Wide Back Notch with Drilling for Hook":
					return UndermountNotch.Wide_Notch;
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
				case "Solid Birch (NO Finger Joint)":
					return MaterialType.SolidBirch;
				case "Solid Birch (No Finger Joint) - SIDES ONLY":
				case "Solid Birch (NO Finger Joint) - SIDES ONLY":
				case "Solid Birch (NO Finger Joint) - ON SIDES ONLY for 7 1/8\", 8 1/ 4\" and 10 1/8\" heights":
					return MaterialType.HybridBirch;
				case "Walnut":
					return MaterialType.Walnut;
				case "1/4\" Bottom":
					return MaterialType.Plywood1_4;
				case "3/8\" Bottom":
					return MaterialType.Plywood3_8;
				case "1/2\" Bottom":
					return MaterialType.Plywood1_2;
				default:
					return MaterialType.Unknown;
			}

		}

		public static RichelieuConfiguration ParseSku(string sku) {

			string specie = sku.Substring(3, 2);
			string botCode = sku.Substring(6, 2);
			string notchCode = sku.Substring(8, 2);
			string frontCode = sku.Substring(10, 1);
			string pullOutCode = sku.Substring(11, 1);
			string rushCode = sku.Substring(sku.Length - 1, 1);

			MaterialType boxMaterial;
			switch (specie) {
				case "08":
					boxMaterial = MaterialType.EconomyBirch;
					break;
				case "09":
					//material = MaterialType.SolidBirch;
					boxMaterial = MaterialType.HybridBirch;
					break;
				default:
					boxMaterial = MaterialType.Unknown;
					break;
			}

			MaterialType botMaterial;
			switch (botCode) {
				case "14":
					botMaterial = MaterialType.Plywood1_4;
					break;
				case "12":
					botMaterial = MaterialType.Plywood1_2;
					break;
				case "38":
					botMaterial = MaterialType.Plywood3_8;
					break;
				default:
					botMaterial = MaterialType.Unknown;
					break;
			}

			UndermountNotch notch;
			switch (notchCode) {
				case "NN":
					notch = UndermountNotch.No_Notch;
					break;
				case "SN":
					notch = UndermountNotch.Std_Notch;
					break;
				case "WN":
					notch = UndermountNotch.Wide_Notch;
					break;
				case "FB":
					notch = UndermountNotch.Front_Back;
					break;
				default:
					notch = UndermountNotch.Unknown;
					break;
			}

			string frontOption = "";
			switch (frontCode) {
				case "H":
					frontOption = "Extra 1\" at top";
					break;
				case "X":
				default:
					frontOption = "";
					break;
            }

			string pullOutOption;
			bool scoopFront = false;
			bool clearFront = false;
			switch (pullOutCode) {
				case "1":
				case "2":
				case "3":
				case "4":
					pullOutOption = "Pull-Out " + pullOutCode;
					scoopFront = true;
					clearFront = true;
					break;
				case "N":
					pullOutOption = "Clear Front";
					clearFront = true;
					break;
				case "R":
				default:
					pullOutOption = "None";
					break;
            }

			bool rush = rushCode.Equals("3");

			return new RichelieuConfiguration {
				BoxMaterial = boxMaterial,
				BotMaterial = botMaterial,
				Notch = notch,
				FrontOption = frontOption,
				Rush = rush,
				ScoopFront = scoopFront,
				PullOutFront = clearFront
			};

		}

		public struct RichelieuConfiguration {
			public MaterialType BoxMaterial { get; set; }
			public MaterialType BotMaterial { get; set; }
			public UndermountNotch Notch { get; set; }
			public string FrontOption { get; set; }
			public bool Rush { get; set; }

			// A scoop front is when the front piece of material is routed out to make a scoop that you can pull out
			public bool ScoopFront { get; set; }
			// A pull out front is a front that is clear birch, even if the box material is economy birch
			public bool PullOutFront { get; set; }
		}

	}

}