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

		private readonly AppSettings _settings;
		public RichelieuExcelDBOrderProvider() {
			_settings = HelperFuncs.ReadSettings();
        }

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
			string orderNote = "";
			foreach (XmlNode linesNode in linesNodes) {
				string description = linesNode.Attributes.GetNamedItem("descriptionEn").InnerText;
				string[] properties = description.Split(',');

				string sku = linesNode.Attributes.GetNamedItem("sku").InnerText;

				RichelieuConfiguration config = ParseSku(sku);
				
				order.Rush = config.Rush;

				string note = linesNode.Attributes.GetNamedItem("note").InnerText;
				if (!string.IsNullOrWhiteSpace(note))
					orderNote += note + "\n";

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
                        ClipsOption = config.Clips,
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

			order.Comment = orderNote;
			System.Windows.Forms.MessageBox.Show(orderNote, "Order Note");

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

		public RichelieuConfiguration ParseSku(string sku) {

			/*
			 * Example: RCT08114ISHNX3R0
			 * 
			 * 0-2		Company Code	|	RCT	
			 * 3-4		MaterialType	|	08->EconomyBirch, 09->Hybrid/Solid Birch, 13->Baltic Birch
			 * 5
			 * 6-7		BottomMaterial	|	12->1/2", 14->1/4", 38->3/8"
			 * 8		Assembly		|	I->Included
			 * 9-10		Notch			|	NN->No Notch, SH->Std Notch, WH->Wide Notch, FB->Front & Back
			 * 11-12	Fasteners		|	NO->Without Fasteners, R4->4 way clips, R6->6 way clips
			 * 13		Front			|	X->Regular, H->Extra 1" at top
			 * 14		Pull-Out		|	R->No Pull, N->Clear Front, 1/2/3->Scoop Front
			 * 15-16	Rush			|	R0->No Rush, R3->3 Day Rush
			 */

			string specie = sku.Substring(3, 2);
			string botCode = sku.Substring(6, 2);
			string notchCode = sku.Substring(9, 2);
			string fastenerCode = sku.Substring(11, 2);
			string frontCode = sku.Substring(13, 1);
			string pullOutCode = sku.Substring(14, 1);
			string rushCode = sku.Substring(15, 2);

			var profile = _settings.MaterialProfiles["richelieu"];
			string boxMaterial;
			string botMaterial;

			try {
				boxMaterial = profile[specie];
				botMaterial = profile[botCode];
			} catch {
				boxMaterial = $"Unknown Material '{specie}'";
				botMaterial = $"Unknown Material '{botCode}'";
			}

			UndermountNotch notch;
			switch (notchCode) {
				case "NN":
					notch = UndermountNotch.No_Notch;
					break;
				case "SH":
					notch = UndermountNotch.Std_Notch;
					break;
				case "WH":
					notch = UndermountNotch.Wide_Notch;
					break;
				case "FB":
					notch = UndermountNotch.Front_Back;
					break;
				default:
					notch = UndermountNotch.Unknown;
					break;
			}

			string clips = "";
			switch (fastenerCode) {
				case "R4":
					clips = "4-Way Richlieu";
					break;
				case "R6":
					clips = "6-Way Richlieu";
					break;
				case "NO":
				default:
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

			bool rush = rushCode.Equals("R3");

			return new RichelieuConfiguration {
				BoxMaterial = boxMaterial,
				BotMaterial = botMaterial,
				Notch = notch,
				FrontOption = frontOption,
				Rush = rush,
				ScoopFront = scoopFront,
				PullOutFront = clearFront,
				Clips = clips
			};

		}

		public struct RichelieuConfiguration {
			public string BoxMaterial { get; set; }
			public string BotMaterial { get; set; }
			public UndermountNotch Notch { get; set; }
			public string FrontOption { get; set; }
			public bool Rush { get; set; }
			public string Clips { get; set; }

			// A scoop front is when the front piece of material is routed out to make a scoop that you can pull out
			public bool ScoopFront { get; set; }
			// A pull out front is a front that is clear birch, even if the box material is economy birch
			public bool PullOutFront { get; set; }

		}

	}

}