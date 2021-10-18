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

namespace ExcelLibTests {

	[TestClass]
	public class RecordRepositoryTests {

		private readonly string _connString = "Data Source=InMemory;Mode=Memory;Cache=Shared";
		private readonly InventoryUseRecord _testRecord;

		public RecordRepositoryTests() {

			_testRecord = new InventoryUseRecord {
				Qty = 1,
				ItemId = 1,
				Timestamp = DateTime.Today,
				JobId = 1
			};

		}

		[TestMethod]
		public void InsertTest() {

			using (var connection = new SqliteConnection(_connString)) {

				InventoryRecordRepository repository = new InventoryRecordRepository(connection);

				connection.Open();
				var newBox = repository.Insert(_testRecord);
				connection.Close();

				Assert.AreEqual(1, newBox.Id);

			}

		}

		[TestMethod]
		public void InsertDeleteTest() {


			using (var connection = new SqliteConnection(_connString)) {

				InventoryRecordRepository repository = new InventoryRecordRepository(connection);

				connection.Open();

				var newBox = repository.Insert(_testRecord);

				repository.Delete(newBox);

				connection.Close();

			}

		}

		[TestMethod]
		public void InsertGetTest() {

			using (var connection = new SqliteConnection(_connString)) {

				InventoryRecordRepository repository = new InventoryRecordRepository(connection);

				connection.Open();

				var newRecord = repository.Insert(_testRecord);

				repository.GetById(newRecord.Id);

				connection.Close();

			}

		}

		[TestMethod]
		public void InsertUpdateTest() {

			using (var connection = new SqliteConnection(_connString)) {

				InventoryRecordRepository repository = new InventoryRecordRepository(connection);

				connection.Open();

				var newRecord = repository.Insert(_testRecord);

				newRecord.Qty = 2;

				repository.Update(newRecord);

				var updatedBox = repository.GetById(newRecord.Id);

				Assert.AreEqual(updatedBox.Qty, updatedBox.Qty);

				connection.Close();

			}

		}


	}
}
