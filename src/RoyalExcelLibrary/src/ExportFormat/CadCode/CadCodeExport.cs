using Newtonsoft.Json;
using RoyalExcelLibrary.ExcelUI.Models;
using RoyalExcelLibrary.ExcelUI.Models.Products;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.ExcelUI.ExportFormat.CadCode {

	public enum PathOffset {
		Left,
		Right,
		Inside,
		Outside,
		Center,
		None
	}

	// Determins whether to mark the part as a 'Small Part' in CADCode
	public enum PartTriState {
		Small,
		Large,
		Default
	}

	public interface IToken {
		object[] GetToken();
	}

	public class Border : IToken {
		public string JobName { get; set; }
		public string ProductId { get; set; }
		public string PartId { get; set; }
		public string PartName { get; set; }
		public int Qty { get; set; }
		public double Width { get; set; }
		public double Height { get; set; }
		public double Thickness { get; set; }
		public string FileName { get; set; }
		public string Face6FileName { get; set; }
		public bool IsFace6 { get; set; }
		public PartTriState PartSize { get; set; }
		public string Material { get; set; }

		public object[] GetToken() {
			return new object[] { JobName, ProductId, PartId, PartName, $"{Qty}", "Border", "", $"{Width}", $"{Height}", $"{Thickness}","", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" ,FileName, "", Face6FileName, IsFace6 ? "Y" : "", "", Material, "", "", "", "", "", "", "", "", "", "", "", PartSize == PartTriState.Small ? "Y" : PartSize == PartTriState.Large ? "N" : ""};
		}
	}

	public class Rectangle : IToken {

		public string Comment { get; set; }
		public double Z_1 { get; set; }
		public double Z_2 { get; set; }
		public double Radius { get; set; }
		public double X_1 { get; set; }
		public double Y_1 { get; set; }
		public double X_2 { get; set; }
		public double Y_2 { get; set; }
		public double X_3 { get; set; }
		public double Y_3{ get; set; }
		public double X_4 { get; set; }
		public double Y_4 { get; set; }
		public PathOffset Offset { get; set; }
		public string Tool { get; set; }
		public int Sequence { get; set; }

		public object[] GetToken() {
			return new object[] { "", "", "", "", "", "Rectangle", "", X_1, Y_1, Z_1, X_2, Y_2, Z_2, X_3, Y_3, X_4, Y_4, Radius, "", "", OffsetString(), "", Tool, "", "", (Sequence == -1 ? "" : $"{Sequence}") };
		}

		public string OffsetString() {
			switch (Offset) {
				case PathOffset.Inside:
					return "I";
				case PathOffset.Outside:
					return "I";
				case PathOffset.Left:
					return "L";
				case PathOffset.Right:
					return "R";
				case PathOffset.Center:
					return "C";
				case PathOffset.None:
				default:
					return "";
			}
		}

	}

	public struct CCPart {
		public Border Border { get; set; }
		public IToken[] Tokens { get; set; }
	}

	public class CadCodeExport {

		private readonly AppSettings _settings;

		public CadCodeExport() {
			_settings = HelperFuncs.ReadSettings();
		}

		public void ExportOrder(Order order, string exportPath) {

			AppSettings settings = HelperFuncs.ReadSettings();

			List<CCPart> parts = new List<CCPart>();

			int startIndex = 1;

			parts.AddRange(CreateUBoxBottoms(order, startIndex));

			startIndex = parts.Count;

			parts.AddRange(CreateTrashTopParts(order, startIndex, settings));

			startIndex = parts.Count;

			parts.AddRange(CreateBottomParts(order, startIndex, settings));

			WriteToFile(exportPath, parts);

		}

		private IEnumerable<CCPart> CreateBottomParts(Order order, int startIndex, AppSettings settings) {

			// Only cutlist standard drawerbox bottoms, ignoring UBoxes
			IEnumerable<DrawerBox> boxes = order.Products
												.Where(p => p.GetType() == typeof(DrawerBox))
												.Cast<DrawerBox>();


			int i = startIndex;
			List<CCPart> parts = new List<CCPart>();
			foreach (DrawerBox box in boxes) {

				double thickness = 0.25 * 25.4;
				if (box.BottomMaterial.Contains("1/2")) {
					thickness = 0.5 * 25.4;
				}

				Border border = new Border {
					JobName = order.Number + " - " + order.Job.Name,
					Width = box.Width - 2 * settings.ManufacturingValues.SideThickness + 2 * settings.ManufacturingValues.DadoDepth - settings.ManufacturingValues.BottomAdj,
					Height = box.Depth - 2 * settings.ManufacturingValues.SideThickness + 2 * settings.ManufacturingValues.DadoDepth - settings.ManufacturingValues.BottomAdj,
					Thickness = thickness,
					Material = box.BottomMaterial,
					ProductId = $"{i}",
					PartId = "Bottom",
					PartName = "Bottom",
					Qty = box.Qty,
					FileName = $"Bottom-{i++}",
					Face6FileName = "",
					PartSize = PartTriState.Default
				};

				parts.Add(new CCPart() {
					Border = border,
					Tokens = new IToken[0]
                });

			}

			return parts;

        }

		private IEnumerable<CCPart> CreateTrashTopParts(Order order, int startIndex, AppSettings settings) {

			IEnumerable<DrawerBox> trashBoxes = order.Products
													.Where(b => b is DrawerBox)
													.Cast<DrawerBox>()
													.Where(b => b.TrashDrawerType != TrashDrawerType.None);

			double canWidth = settings.TrashSettings.CanWidth;
			double canDepth = settings.TrashSettings.CanDepth;
			double singleDepth = settings.TrashSettings.SingleTopMaxDepth;
			double doubleDepth = settings.TrashSettings.DoubleTopMaxDepth;
			double doubleWideDepth = settings.TrashSettings.DoubleWideTopMaxDepth;
			double cutoutRadius = settings.TrashSettings.CutOutRadius;
			double doubleSpaceBetween = settings.TrashSettings.DoubleSpaceBetween;

			int i = startIndex;
			List<CCPart> parts = new List<CCPart>();
			foreach (DrawerBox box in trashBoxes) {

				double boxDepth = box.Depth;

				if (box.TrashDrawerType == TrashDrawerType.Single)
					boxDepth = (singleDepth == -1 || box.Depth <= singleDepth) ? box.Depth : singleDepth;
				else if(box.TrashDrawerType == TrashDrawerType.Double)
					boxDepth = (doubleDepth == -1 || box.Depth <= doubleDepth) ? box.Depth : doubleDepth;
				else if (box.TrashDrawerType == TrashDrawerType.DoubleWide)
					boxDepth = (doubleWideDepth == -1 || box.Depth <= doubleWideDepth) ? box.Depth : doubleWideDepth;

				Border border = new Border {
					JobName = order.Number + " - " + order.Job.Name,
					Width = box.Width,
					Height = boxDepth,
					Thickness = (0.5 * 25.4),
					Material = "1/2\" Ply",
					ProductId = $"{i}",
					PartId = "Trash Top",
					PartName = "Trash Top",
					Qty = box.Qty,  
					FileName = $"TrashTop-{i++}",
					Face6FileName = "",
					PartSize = PartTriState.Default
				};

				List<IToken> tokens = new List<IToken>();

				double centerX = box.Width / 2;
				double centerY = boxDepth / 2;

				switch (box.TrashDrawerType) { 
					case TrashDrawerType.Single:
						tokens.Add(new Rectangle {
							X_1 = centerX - canWidth / 2,
							Y_1 = centerY - canDepth / 2,
							X_3 = centerX + canWidth / 2,
							Y_3 = centerY - canDepth / 2,
							X_2 = centerX + canWidth / 2,
							Y_2 = centerY + canDepth / 2,
							X_4 = centerX - canWidth / 2,
							Y_4 = centerY + canDepth / 2,
							Offset = PathOffset.Inside,
							Tool = "3-8Compt",
							Z_1 = border.Thickness,
							Z_2 = border.Thickness,
							Radius = cutoutRadius
						});
					break;

					case TrashDrawerType.Double:
						double topCenterY = centerY + doubleSpaceBetween/2 + canWidth / 2;
						double bottomCenterY = centerY - doubleSpaceBetween / 2 - canWidth / 2;

						tokens.Add(new Rectangle {
							X_1 = centerX - canDepth / 2,
							Y_1 = topCenterY - canWidth / 2,
							X_3 = centerX + canDepth / 2,
							Y_3 = topCenterY - canWidth / 2,
							X_2 = centerX + canDepth / 2,
							Y_2 = topCenterY + canWidth / 2,
							X_4 = centerX - canDepth / 2,
							Y_4 = topCenterY + canWidth / 2,
							Offset = PathOffset.Inside,
							Tool = "3-8Comp",
							Z_1 = border.Thickness,
							Z_2 = border.Thickness,
							Radius = cutoutRadius
						});

						tokens.Add(new Rectangle {
							X_1 = centerX - canDepth / 2,
							Y_1 = bottomCenterY - canWidth / 2,
							X_3 = centerX + canDepth / 2,
							Y_3 = bottomCenterY - canWidth / 2,
							X_2 = centerX + canDepth / 2,
							Y_2 = bottomCenterY + canWidth / 2,
							X_4 = centerX - canDepth / 2,
							Y_4 = bottomCenterY + canWidth / 2,
							Offset = PathOffset.Inside,
							Tool = "3-8Comp",
							Z_1 = border.Thickness,
							Z_2 = border.Thickness,
							Radius = cutoutRadius
						});
						break;

					case TrashDrawerType.DoubleWide:
						double rightCenterX = centerX + doubleSpaceBetween / 2 + canWidth / 2;
						double leftCenterX = centerX - doubleSpaceBetween / 2 - canWidth / 2;

						tokens.Add(new Rectangle {
							X_1 = rightCenterX - canWidth / 2,
							Y_1 = centerY - canDepth / 2,
							X_3 = rightCenterX + canWidth / 2,
							Y_3 = centerY - canDepth / 2,
							X_2 = rightCenterX + canWidth / 2,
							Y_2 = centerY + canDepth / 2,
							X_4 = rightCenterX - canWidth / 2,
							Y_4 = centerY + canDepth / 2,
							Offset = PathOffset.Inside,
							Tool = "3-8Comp",
							Z_1 = border.Thickness,
							Z_2 = border.Thickness,
							Radius = cutoutRadius
						});

						tokens.Add(new Rectangle {
							X_1 = leftCenterX - canWidth / 2,
							Y_1 = centerY - canDepth / 2,
							X_3 = leftCenterX + canWidth / 2,
							Y_3 = centerY - canDepth / 2,
							X_2 = leftCenterX + canWidth / 2,
							Y_2 = centerY + canDepth / 2,
							X_4 = leftCenterX - canWidth / 2,
							Y_4 = centerY + canDepth / 2,
							Offset = PathOffset.Inside,
							Tool = "3-8Comp",
							Z_1 = border.Thickness,
							Z_2 = border.Thickness,
							Radius = cutoutRadius
						});
						break;

				}

				CCPart part = new CCPart {
					Border = border,
					Tokens = tokens.ToArray()
				};
				parts.Add(part);

			}


			return parts;

        }

		private IEnumerable<CCPart> CreateUBoxBottoms(Order order, int startIndex) {

			IEnumerable<UDrawerBox> uboxes = order.Products
												.Where(b => b is UDrawerBox)
												.Cast<UDrawerBox>();

			int i = startIndex;
			List<CCPart> parts = new List<CCPart>();
			foreach (UDrawerBox box in uboxes) {

				double D1 = box.A - (2 * 16) + (2 * 6) - 1;
				double D2 = box.A - (2 * 16) + (2 * 6) - 1 + box.Width - box.A - box.B + 33 - 1 - (2 * 6) + 1;

				Border border = new Border {
					JobName = order.Number + " - " + order.Job.Name,
					Width = box.Width - (2 * 16) + (2 * 6) - 1,
					Height = box.Depth - (2 * 16) + (2 * 6) - 1,
					Thickness = GetBottomThickness(box.BottomMaterial),
					Material = GetBottomMatCode(box.BottomMaterial),
					ProductId = $"{i}",
					PartId = "UBox Bottom",
					PartName = "UBox Bottom",
					Qty = box.Qty,
					FileName = $"UBottom-{i++}",
					Face6FileName = "",
					PartSize = PartTriState.Default
				};

				Rectangle rectangle = new Rectangle {
					X_1 = D1,
					Y_1 = 0,
					X_2 = D2,
					Y_2 = box.C,
					X_3 = D1,
					Y_3 = box.C,
					X_4 = D2,
					Y_4 = 0,
					Offset = PathOffset.Inside,
					Tool = "3-8Comp",
					Z_1 = border.Thickness,
					Z_2 = border.Thickness
				};

				CCPart part = new CCPart {
					Border = border,
					Tokens = new IToken[] { rectangle }
				};
				parts.Add(part);

			}

			return parts;

		}

		private void WriteToFile(string exportPath, IEnumerable<CCPart> parts) {
			using (FileStream fs = File.Open(exportPath, FileMode.OpenOrCreate)) {

				// Clear contents of file if it already exists
				fs.SetLength(0);

				using (StreamWriter writer = new StreamWriter(fs)) {

					foreach (string heading in header) {
						writer.Write(heading + ",");
						Debug.Write(heading + ",");
					}
					Debug.WriteLine("");
					writer.WriteLine();

					foreach (CCPart part in parts) {
						foreach (object component in part.Border.GetToken()) {
							if (component is null) writer.Write("null,");
							else writer.Write(component.ToString() + ",");
							Debug.Write(component.ToString() + ",");
						}
						Debug.WriteLine("");
						writer.WriteLine();

						foreach (IToken token in part.Tokens) {
							foreach (object component in (token as Rectangle).GetToken()) {
								writer.Write(component.ToString() + ",");
								Debug.Write(component.ToString() + ",");
							}
							Debug.WriteLine("");
							writer.WriteLine();
						}
					}

				}

			}
		}


		private double GetBottomThickness(string material) {
			if (_settings.MaterialThickness.ContainsKey(material))
				return _settings.MaterialThickness[material] * 25.4;
			return 0.25 * 25.4;
		}

		private string GetBottomMatCode(string material) {
			switch (GetBottomThickness(material)) {
				case 0.5*25.4:
					return "Ply-1/2";
				case 0.25 * 25.4:
					return "Ply-1/4";
				default:
					Debug.WriteLine($"UBox has unknown bottom material '{material}'");
					return "UNKNOWN";
			}
		}

		private readonly string[] header = new string[] {
			"JobName",
			"ProductID",
			"PartID",
			"PartName",
			"Quantity",
			"MachiningToken",
			"Face",
			"StartX",
			"StartY",
			"StartZ",
			"EndX",
			"EndY",
			"EndZ",
			"CenterX",
			"CenterY",
			"PocketX",
			"PocketY",
			"Radius",
			"Pitch",
			"-",
			"OffsetSide",
			"-",
			"ToolNumber",
			"ToolDiameter",
			"-",
			"SequenceNumber",
			"-",
			"Filename",
			"-",
			"Face 6 Filename",
			"Face 6 Flag",
			"-",
			"Material",
			"Graining",
			"-",
			"Rotation",
			"",
			"Arc Direction",
			"Start Angle",
			"End Angle",
			"",
			"Feed Speed",
			"Spindle Speed",
			"-",
			"Small Part",
			};

	}

}
