using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using RoyalExcelLibrary.ExcelUI.Models;

using Excel = Microsoft.Office.Interop.Excel;

namespace RoyalExcelLibrary.ExcelUI.Providers {
	public class InventoryProvider {

		private readonly Excel.Worksheet _worksheet;

		public InventoryProvider(Excel.Worksheet worksheet) {
			_worksheet = worksheet;
		}

		public IEnumerable<InventoryItem> LoadAvailableInventory() {

			List<InventoryItem> availableItems = new List<InventoryItem>();

			double thickness = 15.875;

			Excel.Range header = _worksheet.Range["AvailableInventory"];

			int maxItems = 200;
			for (int i = 0; i < maxItems; i++) {

				Excel.Range row;
				try {
					row = header.Offset[i + 1];
				} catch {
					System.Windows.Forms.MessageBox.Show($"Unable to read data");
					break;
				}

				try {
					string typeStr = row.Item[1, 1].Text;

					if (string.IsNullOrEmpty(typeStr)) break;

					double width = row.Item[1, 2].Value2;
					double length = row.Item[1, 3].Value2;

					InventoryItem item = new InventoryItem();
					item.Length = length;
					item.Width = width;
					item.Thickness = thickness;
					item.Name = $"{typeStr}-{width}x{length}";
					item.Material = MaterialFunctions.StringToType(typeStr);

					availableItems.Add(item);

				} catch (Exception e) {
					Debug.WriteLine(e);
					System.Windows.Forms.MessageBox.Show($"line #{i} is invalid");
				}

			}

			return availableItems;

		}

	}
}
