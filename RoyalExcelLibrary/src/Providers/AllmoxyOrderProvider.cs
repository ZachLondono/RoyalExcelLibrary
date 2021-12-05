using RoyalExcelLibrary.Models;
using System;
using System.Xml;
using System.Collections.Generic;
using System.Diagnostics;
using RoyalExcelLibrary.Models.Products;
using RoyalExcelLibrary.Models.Options;
using RoyalExcelLibrary.ExportFormat;

namespace RoyalExcelLibrary.Providers {

	public class AllmoxyOrderProvider : IFileOrderProvider {

		public string FilePath { get; set; }

		private XmlNode _currentOrderNode;
		private bool _isDocLoaded;
		private int _orderNum;

		public AllmoxyOrderProvider() {
			_isDocLoaded = false;
			_orderNum = 1;
		}

		private void LoadFile() {

			if (_isDocLoaded) return;

			XmlDocument doc = new XmlDocument();

			if (string.IsNullOrEmpty(FilePath))
				throw new InvalidOperationException("No file path set");

			doc.Load(FilePath);

			// TODO IF DATA IS A BATCH OF JOBS, THE ROOT ELEMENT WILL BE BATCH NOT ORDER
			_currentOrderNode = doc.FirstChild;
			if (_currentOrderNode.LocalName.Equals("xml")) {
				_currentOrderNode = _currentOrderNode.NextSibling;
			}

		}

		public bool MoveToNextOrder() {
			_currentOrderNode = _currentOrderNode.NextSibling;
			if (_currentOrderNode is null) return false;
			_orderNum++;
			return true;
		}

		public Order LoadCurrentOrder() {

			LoadFile();
			
			string id_str = _currentOrderNode.Attributes.GetNamedItem("id")?.Value ?? "---";
						
			XmlElement xmlElement = _currentOrderNode as XmlElement;

			string customer = xmlElement["customer"].InnerText;
			string name = xmlElement["name"].InnerText;
			string date = xmlElement["date"].InnerText;
			string total = xmlElement["total"].InnerText;

			var shipping = _currentOrderNode.SelectSingleNode($"/order[{_orderNum}]/shipping");
			var shipMethod = shipping["method"]?.InnerText ?? "";
			Address shippingAddress = null;
			if (!shipMethod.Equals("Pickup")) {
				try {
					string shipAddress = shipping["address"]?.InnerText ?? "";
					var addressParts = shipAddress.Split(',');

					string streetAddress1 = addressParts[1].Trim();
					string streetAddress2 = "";
					if (addressParts.Length > 5) streetAddress2 = addressParts[2].Trim();
					string city = addressParts[addressParts.Length - 3].Trim();
					string state_zip = addressParts[addressParts.Length - 2].Trim();
					var arr = state_zip.Split(' ');
					string state = arr[0].Trim(); // state_zip has a preceding space
					string zip = arr[1].Trim();
					string country = addressParts[addressParts.Length - 1].Trim();

					shippingAddress = new Address {
						Line1 = streetAddress1,
						Line2 = streetAddress2,
						City = city,
						State = state,
						Zip = zip,
					};

				} catch {
					Debug.WriteLine("Error reading shipping address");
				}
			} else {
				shippingAddress = new Address {
					Line1 = "Pickup",
					Line2 = "",
					City = "",
					State = "",
					Zip = ""
				};
			}

			var invoice = _currentOrderNode.SelectSingleNode($"/order[{_orderNum}]/invoice");
			decimal subtotal = Convert.ToDecimal(invoice["subtotal"]?.InnerText ?? "0");
			decimal tax = Convert.ToDecimal(invoice["tax"]?.InnerText ?? "0");
			decimal shippingPrice = Convert.ToDecimal(invoice["shipping"]?.InnerText ?? "0");
			
			var drawerboxes = _currentOrderNode.SelectNodes($"/order[{_orderNum}]/DrawerBox");	//TODO: get only the drawer boxes in the current order (if batch order)

			List<DrawerBox> boxes = new List<DrawerBox>();

			int lineNum = 1;
			foreach (XmlNode drawerbox in drawerboxes) {

				DrawerBox box;

				XmlNode dimensions = drawerbox["dimensions"];
				XmlNode udimensions = drawerbox["udimensions"];

				if (udimensions is null) {
                    box = new DrawerBox {
                        ProductName = "Standard Drawer Box"
                    };
                } else {

					double a = HelperFuncs.ConvertToDouble(udimensions["a"].InnerText);
					double b = HelperFuncs.ConvertToDouble(udimensions["b"].InnerText);
					double c = HelperFuncs.ConvertToDouble(udimensions["c"].InnerText);

					box = new UDrawerBox() {
						A = a * 25.4,
						B = b * 25.4,
						C = c * 25.4
					};
					box.ProductName = "UShaped Drawer Box";
				}

				box.ProductName = "Drawer Box";

				double height = HelperFuncs.ConvertToDouble(dimensions["height"].InnerText);
				double width = HelperFuncs.ConvertToDouble(dimensions["width"].InnerText);
				double depth = HelperFuncs.ConvertToDouble(dimensions["depth"].InnerText);
				int qty = Convert.ToInt32(drawerbox["qty"].InnerText);


                MaterialType sideMaterial = MapMaterial(drawerbox["material"].InnerText, out bool postfinish);
                MaterialType bottomMaterial = MapMaterial(drawerbox["bottom"].InnerText, out bool throwaway);
                string insert = drawerbox["insert"]?.InnerText ?? "";
				UndermountNotch notch = MapNotch(drawerbox["notch"]?.InnerText ?? "");
				Clips clips = MapClips(drawerbox["clips"]?.InnerText ?? "");
				bool logo = drawerbox["logo"].InnerText.Equals("Yes");
				bool scoop = drawerbox["scoop"].InnerText.Equals("Yes");
				string labelNote = drawerbox["note"]?.InnerText ?? "";
				Decimal unitPrice = Convert.ToDecimal(drawerbox["price"]?.InnerText ?? "0");

				box.SideMaterial = sideMaterial;
				box.BottomMaterial = bottomMaterial;
				box.Height = height * 25.4;
				box.Width = width * 25.4;
				box.Depth = depth * 25.4;
				box.Qty = qty;
				box.Note = labelNote;
				box.ClipsOption = clips;
				box.InsertOption = insert;
				box.NotchOption = notch;
				box.ScoopFront = scoop;
				box.MountingHoles = false;
				box.PostFinish = postfinish;
				box.Logo = logo;
				box.UnitPrice = unitPrice;
				box.LineNumber = lineNum++;

				boxes.Add(box);

			}

			Job job = new Job {
				JobSource = "Allmoxy",
				CreationDate = string.IsNullOrEmpty(date) ? DateTime.Today :  DateTime.Parse(date),
				GrossRevenue = Convert.ToDecimal(total) * 0.87M,
				Name = name
			};

			Order order = new Order(job);
			order.AddProducts(boxes);
			order.Number = id_str;
			order.SubTotal = subtotal;
			order.Tax = tax;
			order.ShippingCost = shippingPrice;
			order.Customer = new Company {
				Name = customer,
				Address = shippingAddress
			};
			
			Address royalAddress = new Address {
				Line1 = "15E Easy St",
				Line2 = "",
				City = "Bound Brook",
				State = "NJ",
				Zip = "08805"
			};

			order.Vendor = new Company {
				Name = "OT",
				Address = royalAddress
			};

			order.Supplier = new Company {
				Name = "Metro Cabinet Parts",
				Address = royalAddress
			};

			return order;
		}

		private MaterialType MapMaterial(string text, out bool post_finish) {

			post_finish = false;

			switch (text) {
				case "1/4\" Plywood":
					return MaterialType.Plywood1_4;
				case "1/2\" Plywood":
					return MaterialType.Plywood1_2;
				case "Post-Finished Birch":
					post_finish = true;
					return MaterialType.SolidBirch;
				case "Pre-Finished Birch":
					return MaterialType.SolidBirch;
				case "Walnut":
					post_finish = true;
					return MaterialType.SolidWalnut;
				case "Walnut - Unfinished":
					return MaterialType.SolidWalnut;
				case "White Oak":
					post_finish = true;
					return MaterialType.WhiteOak;
				case "White Oak - Unfinished":
					return MaterialType.WhiteOak;
				default:
					return MaterialType.Unknown;
			}

		}

		private Clips MapClips(string text) {
			switch (text) {
				case "Blum":
					return Clips.Blum;
				case "Hettich":
					return Clips.Hettich;
				case "Richelieu":
					return Clips.Richelieu;
				case "":
				case "No Clips":
					return Clips.No_Clips;
				default:
					return Clips.Unknown;
			}
		}

		private UndermountNotch MapNotch(string text) {
			switch (text) {
				case "Notch for Standard U/M Slide":
					return UndermountNotch.Std_Notch;
				case "Notch for U/M Slide, Wide":
					return UndermountNotch.Wide_Notch;
				case "Notch for 828":
					return UndermountNotch.Notch_828;
				case "":
				case "No Notch":
					return UndermountNotch.No_Notch;
				default:
					return UndermountNotch.Unknown;
			}
		}

	}

}
