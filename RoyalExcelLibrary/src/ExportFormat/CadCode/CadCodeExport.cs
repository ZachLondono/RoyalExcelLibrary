using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.ExportFormat.CadCode {

	public enum PathOffset {
		Left,
		Right,
		Inside,
		Outside,
		Center,
		None
	}

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

		public void ExportOrder(Order order, string exportPath) {

			IEnumerable<UDrawerBox> uboxes = order.Products	
												.Where(b => b is UDrawerBox)
												.Cast<UDrawerBox>();

			int i = 1;
			List<CCPart> parts = new List<CCPart>();
			foreach (UDrawerBox box in uboxes) {

				double D1 = box.A - (2 * 16) + (2 * 6) - 1;
				double D2 = box.A - (2 * 16) + (2 * 6) - 1 + box.Width - box.A - box.B + 33 - 1 - (2*6) + 1;

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
                    Tool = "212",
                    Z_1 = border.Thickness,
                    Z_2 = border.Thickness
                };

                CCPart part = new CCPart {
                    Border = border,
                    Tokens = new IToken[] { rectangle }
                };
                parts.Add(part);

			}

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

						foreach (IToken token in part.Tokens){
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

		private double GetBottomThickness(MaterialType material) {
			switch (material) {
				case MaterialType.Plywood1_2:
				case MaterialType.BlackMela1_2:
				case MaterialType.WhiteMela1_2:
					return 25.4 / 2;
				case MaterialType.Plywood1_4:
				case MaterialType.BlackMela1_4:
				case MaterialType.WhiteMela1_4:
					return 25.4 / 4;
				default:
					Debug.WriteLine($"UBox has unknown bottom material '{material}'");
					return 0;
			}
		}

		private string GetBottomMatCode(MaterialType material) {
			switch (material) {
				case MaterialType.Plywood1_2:
				case MaterialType.BlackMela1_2:
				case MaterialType.WhiteMela1_2:
					return "Ply-1/2";
				case MaterialType.Plywood1_4:
				case MaterialType.BlackMela1_4:
				case MaterialType.WhiteMela1_4:
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
