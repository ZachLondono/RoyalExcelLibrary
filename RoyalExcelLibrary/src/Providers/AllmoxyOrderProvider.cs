using RoyalExcelLibrary.Models;
using System;
using System.Xml;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using RoyalExcelLibrary.Models.Products;
using RoyalExcelLibrary.Models.Options;
using RoyalExcelLibrary.ExportFormat;

namespace RoyalExcelLibrary.Providers {

	public class AllmoxyOrderProvider : IOrderProvider {

		private readonly string _importPath;
		private XmlDocument _document;
		private XmlNode _currentOrderNode;
		private bool _isDocLoaded;
		private int _orderNum;

		public AllmoxyOrderProvider(string importPath) {
			_importPath = importPath;
			_isDocLoaded = false;
			_orderNum = 1;
		}

		private void LoadFile() {

			if (_isDocLoaded) return;

			XmlDocument doc = new XmlDocument();

			doc.Load(_importPath);

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
			
			string id_str = "A" + _currentOrderNode.Attributes.GetNamedItem("id")?.Value ?? "---";
						
			XmlElement xmlElement = _currentOrderNode as XmlElement;

			string customer = xmlElement["customer"].InnerText;
			string name = xmlElement["name"].InnerText;
			string date = xmlElement["date"].InnerText;
			string description = xmlElement["description"].InnerText;
			string status = xmlElement["status"].InnerText;
			string total = xmlElement["total"].InnerText;

			var shipping = _currentOrderNode.SelectSingleNode($"/order[{_orderNum}]/shipping");
			var shipMethod = shipping["method"]?.InnerText ?? "";
			Address shippingAddress = null;
			if (!shipMethod.Equals("Pickup")) {
				try {
					string shipAddress = shipping["address"]?.InnerText ?? "";
					var addressParts = shipAddress.Split(',');

					string streetAddress = addressParts[1];
					string city = addressParts[addressParts.Length - 3];
					string state_zip = addressParts[addressParts.Length - 2];
					var arr = state_zip.Split(' ');
					string state = arr[1]; // state_zip has a preceding space
					string zip = arr[2];
					string country = addressParts[addressParts.Length - 1];

					shippingAddress = new Address {
						StreetAddress = streetAddress,
						City = city,
						State = state,
						Zip = zip,
					};

				} catch {
					Debug.WriteLine("Error reading shipping address");
				}
			}

			var drawerboxes = _currentOrderNode.SelectNodes($"/order[{_orderNum}]/DrawerBox");	//TODO: get only the drawer boxes in the current order (if batch order)

			List<DrawerBox> boxes = new List<DrawerBox>();

			foreach (XmlNode drawerbox in drawerboxes) {

				DrawerBox box;

				XmlNode dimensions = drawerbox["dimensions"];
				XmlNode udimensions = drawerbox["udimensions"];

				if (udimensions is null) {
					box = new DrawerBox();
				} else {

					double a = ConvertToDouble(udimensions["a"].InnerText);
					double b = ConvertToDouble(udimensions["b"].InnerText);
					double c = ConvertToDouble(udimensions["c"].InnerText);

					box = new UDrawerBox() {
						A = a * 25.4,
						B = b * 25.4,
						C = c * 25.4
					};
				}


				double height = ConvertToDouble(dimensions["height"].InnerText);
				double width = ConvertToDouble(dimensions["width"].InnerText);
				double depth = ConvertToDouble(dimensions["depth"].InnerText);
				int qty = Convert.ToInt32(drawerbox["qty"].InnerText);


				bool postfinish;
				MaterialType sideMaterial = MapMaterial(drawerbox["material"].InnerText, out postfinish);
				bool throwaway;
				MaterialType bottomMaterial = MapMaterial(drawerbox["bottom"].InnerText, out throwaway);
				Insert insert = MapInsert(drawerbox["insert"]?.InnerText ?? "");
				UndermountNotch notch = MapNotch(drawerbox["notch"]?.InnerText ?? "");
				Clips clips = MapClips(drawerbox["clips"]?.InnerText ?? "");
				bool logo = drawerbox["logo"].InnerText.Equals("Yes");
				bool scoop = drawerbox["scoop"].InnerText.Equals("Yes");
				string labelNote = drawerbox["note"]?.InnerText ?? "";
				double unitPrice = Convert.ToDouble(drawerbox["price"]?.InnerText ?? "0");

				box.SideMaterial = sideMaterial;
				box.BottomMaterial = bottomMaterial;
				box.Height = height * 25.4;
				box.Width = width * 25.4;
				box.Depth = depth * 25.4;
				box.Qty = qty;
				box.LabelNote = labelNote;
				box.ClipsOption = clips;
				box.InsertOption = insert;
				box.NotchOption = notch;
				box.ScoopFront = scoop;
				box.MountingHoles = false;
				box.PostFinish = postfinish;
				box.Logo = logo;
				box.UnitPrice = unitPrice;

				boxes.Add(box);

			}

			Job job = new Job {
				JobSource = "Allmoxy",
				Status = Status.Confirmed,
				CreationDate = string.IsNullOrEmpty(date) ? DateTime.Today :  DateTime.Parse(date),
				GrossRevenue = Convert.ToDouble(total) * 0.87,
				Name = name
			};

			Order order = new Order(job, customer, id_str);
			order.AddProducts(boxes);
			order.ShipAddress = shippingAddress;

			return order;
		}

		private MaterialType MapMaterial(string text, out bool post_finish) {

			post_finish = false;

			switch (text) {
				case "1/4\" Plywood":
					return MaterialType.Plywood1_4;
				case "1/2\" Plywood":
					return MaterialType.Plywood1_4;
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
				default:
					return MaterialType.Unknown;
			}

		}

		private Insert MapInsert(string text) {
			switch (text) {
				case "Cutlery Inser 15\"":
					return Insert.Cutlery_15;
				case "Cutlery Inser 23 1/2\"":
					return Insert.Cutlery_23;
				case "Fixed Divider 2":
					return Insert.Divider_2;
				case "Fixed Divider 3":
					return Insert.Divider_3;
				case "Fixed Divider 4":
					return Insert.Divider_4;
				case "Fixed Divider 5":
					return Insert.Divider_5;
				case "Fixed Divider 6":
					return Insert.Divider_6;
				case "Fixed Divider 7":
					return Insert.Divider_7;
				case "Fixed Divider 8":
					return Insert.Divider_8;
				case "":
				case "None":
					return Insert.No_Insert;
				default:
					return Insert.Unknown;
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

		// <summary>Converts a string into a double</summary>
		// <remark>
		// Attempts to use the Convert.ToDouble method, however if the string is a fraction it will do the conversion by splitting the number up into it's whole number, numerator and denominator sections and converting each to a double
		// </remark>
		private double ConvertToDouble(string text) {

			try {
				return Convert.ToDouble(text);
			} catch (FormatException) {

				string[] parts = text.Split(' ', '/');

				double val = Convert.ToDouble(parts[0]);
				if (parts.Length == 3) {

					double numerator = Convert.ToDouble(parts[1]);
					double denomenator = Convert.ToDouble(parts[2]);

					val += numerator / denomenator;

				}

				return val;

			}

		}

	}

}
