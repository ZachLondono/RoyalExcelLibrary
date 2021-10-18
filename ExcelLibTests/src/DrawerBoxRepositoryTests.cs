using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.DAL.Repositories;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Data.Sqlite;
using System.Diagnostics;
using RoyalExcelLibrary.Models.Products;

namespace ExcelLibTests{
	
	[TestClass]
	public class DrawerBoxRepositoryTests {

		private readonly string _connString = "Data Source=InMemory;Mode=Memory;Cache=Shared";
		private readonly DrawerBox _testBox;

		public DrawerBoxRepositoryTests() {

			_testBox = new DrawerBox {
				Qty = 1,
				Height = 105,
				Width = 500,
				Depth = 500,
				SideMaterial = MaterialType.Unknown,
				BottomMaterial = MaterialType.Unknown,
				JobId = 1
			};

		}

		[TestMethod]
		public void InsertTest() {

			using (var connection = new SqliteConnection(_connString)) {

				DrawerBoxRepository repository = new DrawerBoxRepository(connection);

				connection.Open();
				var newBox = repository.Insert(_testBox);
				connection.Close();

				Assert.AreEqual(1, newBox.Id);

			}

		}

		[TestMethod]
		public void InsertDeleteTest() {


			using (var connection = new SqliteConnection(_connString)) {

				DrawerBoxRepository repository = new DrawerBoxRepository(connection);

				connection.Open();
				
				var newBox = repository.Insert(_testBox);

				repository.Delete(newBox);

				connection.Close();

			}

		}

		[TestMethod]
		public void InsertGetTest() {

			using (var connection = new SqliteConnection(_connString)) {

				DrawerBoxRepository repository = new DrawerBoxRepository(connection);

				connection.Open();

				var newBox = repository.Insert(_testBox);

				repository.GetById(newBox.Id);

				connection.Close();

			}

		}

		[TestMethod]
		public void InsertUpdateTest() {

			using (var connection = new SqliteConnection(_connString)) {

				DrawerBoxRepository repository = new DrawerBoxRepository(connection);

				connection.Open();

				var newBox = repository.Insert(_testBox);

				newBox.Height = 123;

				repository.Update(newBox);

				var updatedBox = repository.GetById(newBox.Id);

				Assert.AreEqual(newBox.Height, updatedBox.Height);

				connection.Close();

			}

		}


	}
}
