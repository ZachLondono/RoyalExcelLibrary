using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Xml;

namespace RoyalExcelLibrary.Services {

	// </summary>
	// This class serves as a wrapper class over the old Dymo COM SDK.
	// </summary>
	// <remark>
	// This specific version of the dymo sdk was used because it was the only version that would work with our label printer. 
	// This sdk version is only compatible with the older .label format labels. These label templates can only be created with the Dymo v8 software, not the newer dymo connect.
	// The labels use the xml node <DieCutLabel> rather than the new format which uses <DesktopLabel>
	// </remark>
	public class DymoLabelService : ILabelService {

		private readonly string _labelFile;
		private readonly Dictionary<string, LabelField> _labelFields;
		private readonly Dictionary<Label, int> _labels;

		public DymoLabelService(string labelFile) {

			_labelFile = labelFile;
			XmlDocument doc = new XmlDocument();
			doc.Load(labelFile);
			var labelObjectNodes = doc.SelectNodes("/DieCutLabel/ObjectInfo");
			if (labelObjectNodes is null)
				throw new ArgumentException($"The provided file is not a valid label template\n{_labelFile}");

			// A dictionary mapping the textobject's name to the LabelField instance, which holds its value and type
			_labelFields = new Dictionary<string, LabelField>();

			// Read the label file to find all the fillable text objects in the label
			// Each of theses text objects can have their values changed
			foreach (XmlNode labelObjectInfo in labelObjectNodes) {
				XmlNodeList childObject = labelObjectInfo.ChildNodes;
				foreach (XmlNode labelObjectNode in childObject) {
					if (labelObjectNode.Name.Equals("TextObject")) {
						_labelFields.Add(labelObjectNode["Name"].InnerText, new LabelField {
							Type = LabelFieldType.Text,
							Value = ""
						});
						break;
					} else if (labelObjectNode.Name.Equals("AddressObject")) {
						_labelFields.Add(labelObjectNode["Name"].InnerText, new LabelField {
							Type = LabelFieldType.Address,
							Value = ""
						});
						break;
					}
				}
			}

			_labels = new Dictionary<Label, int>();

		}

		// <summary>
		// Returns a new label which has all the same label fields as the current template 
		// </summary>
		public Label CreateLabel() {
			// Each label must have it's own unique field mapping so that they can each have their own values in each field
			Dictionary<string, LabelField> newFields = new Dictionary<string, LabelField>();
			// making a simple copy of the dictionary would still have each value reference the same LabelField instances, so instead we create a new LabelField instance for each field
			foreach (var field in _labelFields)
				newFields.Add(field.Key, new LabelField { Type = field.Value.Type, Value = null });

            Label label = new Label {
                LabelFields = newFields
            };
            return label;
		}

		public void AddLabel(Label label, int qty) {
			_labels.Add(label, qty);
		}

		public void PrintLabels() {
			
			var dymo = new DYMO.DLS.SDK.DLS7SDK.DymoHighLevelSDK();
			var addin = dymo.DymoAddin;

			string printerName = addin.GetCurrentPrinterName();
#if !DEBUG
			if (!addin.IsPrinterOnline(printerName))
				throw new ArgumentException($"Printer '{addin.GetCurrentPrinterName()}' is not online");
#endif

			var selected = addin.SelectPrinter(printerName);
#if !DEBUG
			if (!selected)
				throw new ArgumentException($"Not able to select Dymo printer for printing\nPrinter '{addin.GetCurrentPrinterName()}' is not available");
#endif

			int i = 0;
			foreach (Label label in _labels.Keys) {
				
				var loaded = addin.Open(_labelFile);
				if (!loaded)
					throw new InvalidOperationException($"The label template file is not valid\n'{_labelFile}'");

				var d_label = dymo.DymoLabels;

				foreach (var item in label.LabelFields) {
					string objectName = item.Key;
					LabelField field = item.Value;
					if (field.Value is null) continue;

					d_label.SetField(objectName, field.Value.ToString());
				}

#if DEBUG
				addin.SaveAs($@"C:\Users\Zachary Londono\Desktop\label-{i++}.label");
#else
				addin.Print(_labels[label], false);
#endif
			}

		}
	}

}