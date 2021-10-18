using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Linq;

using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using RoyalExcelLibrary.Services;
using RoyalExcelLibrary.DAL;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Data.Sqlite;

namespace ExcelLibTests {

	[TestClass]
	public class InventoryUseTests {

		private readonly Order _order;
		private readonly IEnumerable<IProduct> _solid_products;
		private readonly IEnumerable<IProduct> _economy_products;
		private readonly IEnumerable<IProduct> _large_products;
		private readonly IEnumerable<InventoryItem> _availableInventory;

		public InventoryUseTests() {

			DrawerBox box = new DrawerBox();
			box.Height = 10;
			box.Width = 10;
			box.Depth = 10;
			box.Qty = 1;
			box.SideMaterial = MaterialType.SolidBirch;
			box.BottomMaterial = MaterialType.Plywood1_2;

			List<IProduct> solid_products = new List<IProduct>();
			solid_products.Add(box);

			_solid_products = solid_products;


			DrawerBox box2 = new DrawerBox();
			box2.Height = 10;
			box2.Width = 10;
			box2.Depth = 10;
			box2.Qty = 1;
			box2.SideMaterial = MaterialType.EconomyBirch;
			box2.BottomMaterial = MaterialType.Plywood1_2;

			List<IProduct> economy_products = new List<IProduct>();
			economy_products.Add(box2);

			_economy_products = economy_products;


			DrawerBox box3 = new DrawerBox();
			box3.Height = 10;
			box3.Width = 20;
			box3.Depth = 20;
			box3.Qty = 1;
			box3.SideMaterial = MaterialType.EconomyBirch;
			box3.BottomMaterial = MaterialType.Plywood1_2;

			List<IProduct> large_products = new List<IProduct>();
			large_products.Add(box3);

			_large_products = large_products;

			List<InventoryItem> availableInventory = new List<InventoryItem>();

			InventoryItem item = new InventoryItem();
			item.Name = "solid";
			item.Id = 1;
			item.IsAvailable = true;
			item.Width = 10;
			item.Length = 100;
			item.Thickness = 15.9;
			item.Material = MaterialType.SolidBirch;

			InventoryItem item2 = new InventoryItem();
			item2.Name = "economy";
			item2.Id = 2;
			item2.IsAvailable = true;
			item2.Width = 10;
			item2.Length = 40;
			item2.Thickness = 15.9;
			item2.Material = MaterialType.EconomyBirch;

			availableInventory.Add(item);
			availableInventory.Add(item2);

			_availableInventory = availableInventory;

			_order = new Order(new Job {
				Name = "ABC",
				Id = 0,
				CreationDate = DateTime.Now
			}); ;

			_order.AddProduct(box);
			_order.AddProduct(box2);
			_order.AddProduct(box3);

		}

		// <summary>
		// Testing a product which should all be able to be cut from the same part
		// </summary>
		[TestMethod]
		public void CalculateOptimizedPartsSolid() {
			IEnumerable<Part> unplacedParts;
			IEnumerable<InventoryUseRecord> itemsNeeded = InventoryService.GetOptimizedParts(_availableInventory, _solid_products, 1, 0, out unplacedParts);
			Assert.AreEqual(1, itemsNeeded.Count());
			Debug.WriteLine("Item ID: " + itemsNeeded.FirstOrDefault().ItemId);
		}

		[TestMethod]
		public void CalculateOptimizedPartsEconomy() {
			IEnumerable<Part> unplacedParts;
			IEnumerable<InventoryUseRecord> itemsNeeded = InventoryService.GetOptimizedParts(_availableInventory, _economy_products, 1, 0, out unplacedParts);
			Assert.AreEqual(1, itemsNeeded.Count());
			Debug.WriteLine("Item ID: " + itemsNeeded.FirstOrDefault().ItemId);
		}
		
		// <summary>
		// Product is too large to be cut out of just one piece
		// </summary>
		[TestMethod]
		public void CalculateOptimizedPartsLarge() {
			IEnumerable<Part> unplacedParts;
			IEnumerable<InventoryUseRecord> itemsNeeded = InventoryService.GetOptimizedParts(_availableInventory, _large_products, 1, 0, out unplacedParts);
			Assert.AreEqual(1, itemsNeeded.Count());
			Assert.AreEqual(2, itemsNeeded.FirstOrDefault().Qty);
			Debug.WriteLine("Item ID: " + itemsNeeded.FirstOrDefault().ItemId);
		}

		// <summary>
		// Testing having significant waste, which will lead to more items required
		// </summary>
		[TestMethod]
		public void CalculateOptimizedPartsWithWaste() {
			IEnumerable<Part> unplacedParts;
			IEnumerable<InventoryUseRecord> itemsNeeded = InventoryService.GetOptimizedParts(_availableInventory, _economy_products, 1, 1, out unplacedParts);
			Assert.AreEqual(1, itemsNeeded.Count());
			Assert.AreEqual(2, itemsNeeded.FirstOrDefault().Qty);
			Debug.WriteLine("Item ID: " +itemsNeeded.FirstOrDefault().ItemId);
		}

		[TestMethod]
		public void TrackOrderMaterial() {

			using (SqliteConnection connection = new SqliteConnection("Data Source=InMemory;Mode=Memory;Cache=Shared")) {

				InventoryService service = new InventoryService(connection);
				IEnumerable<Part> unplacedParts;

				connection.Open();
				service.TrackMaterialUsage(_order, out unplacedParts);
				connection.Close();

			}

		}

	}

}
