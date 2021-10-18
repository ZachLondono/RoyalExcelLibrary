using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RoyalExcelLibrary.Extensions;
using RoyalExcelLibrary.Models;

namespace RoyalExcelLibrary.DAL.Repositories {
	public class InventoryRecordRepository : IInventoryRecordRepository {

		private readonly string _invRecordTable = "inventory_use";
		private readonly string _idCol = "id";
		private readonly string _qtyCol = "qty";
		private readonly string _itemCol = "item_id";
		private readonly string _jobCol = "job_id";
		private readonly string _dateCol = "timestamp";

		private bool isTableCreated = false;
		private readonly IDbConnection _connection;

		public InventoryRecordRepository(IDbConnection connection) {
			_connection = connection;
		}

		public void Delete(InventoryUseRecord entity) {

			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $@"DELETE FROM {_invRecordTable} 
									WHERE {_idCol} = @id;";

			command.AddParamWithValue("@id", entity.Id);

			try {
				if (command.ExecuteNonQuery() != 1) throw new InvalidOperationException("Unable to delete entity");
			} catch (Exception e) {
				Debug.WriteLine(e);
			}


		}

		public InventoryUseRecord GetById(int id) {

			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $@"SELECT {_qtyCol}, {_itemCol}, {_jobCol}, {_dateCol}	
									FROM {_invRecordTable}
									WHERE {_idCol} = @id;";

			command.AddParamWithValue("@id", id);

			using (var reader = command.ExecuteReader()) {

				reader.Read();

				var itemRecord = new InventoryUseRecord {
					Qty = reader.GetInt32(0),
					ItemId = reader.GetInt32(1),
					JobId = reader.GetInt32(2),
					Timestamp = reader.GetDateTime(3),
					Id = id
				};

				return itemRecord;

			}

			throw new InvalidOperationException($"Entity with id '{id}' does not exist");

		}

		public InventoryUseRecord Insert(InventoryUseRecord entity) {

			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $"INSERT INTO {_invRecordTable} ({_qtyCol}, {_itemCol}, {_jobCol}, {_dateCol}) VALUES (@qty, @item, @job, @timestamp); SELECT last_insert_rowid();";
			command.AddParamWithValue("@qty", entity.Qty);
			command.AddParamWithValue("@item", entity.ItemId);
			command.AddParamWithValue("@job", entity.JobId);
			command.AddParamWithValue("@timestamp", entity.Timestamp);

			using (var reader = command.ExecuteReader()) {

				if (!reader.Read()) {
					Debug.WriteLine("noting returned");
					return null;
				}

				int newId = reader.GetInt32(0);

				entity.Id = newId;

			}

			return entity;

		}

		public void Update(InventoryUseRecord entity) {
			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $"UPDATE {_invRecordTable} SET {_qtyCol} = @qty, {_itemCol} = @item, {_jobCol} = @job, {_dateCol} = @timestamp WHERE {_idCol} = @id;";
			command.AddParamWithValue("@id", entity.Id);
			command.AddParamWithValue("@qty", entity.Qty);
			command.AddParamWithValue("@item", entity.ItemId);
			command.AddParamWithValue("@job", entity.JobId);
			command.AddParamWithValue("@timestamp", entity.Timestamp);

			if (command.ExecuteNonQuery() != 1) throw new InvalidOperationException("Unable to update record");

		}

		public IEnumerable<InventoryUseRecord> GetAll() {

			var command = _connection.CreateCommand();
			command.CommandText = $@"SELECT {_qtyCol}, {_itemCol}, {_jobCol}, {_dateCol}, {_idCol}
									FROM {_invRecordTable};";

			List<InventoryUseRecord> records = new List<InventoryUseRecord>();

			using (var reader = command.ExecuteReader()) {

				while (reader.Read()) {

					var itemRecord = new InventoryUseRecord {
						Qty = reader.GetInt32(0),
						ItemId = reader.GetInt32(1),
						JobId = reader.GetInt32(2),
						Timestamp = reader.GetDateTime(3),
						Id = reader.GetInt32(4)
					};

					records.Add(itemRecord);

				}

			}

			return records;

		}

		private void CreateTable() {
			if (isTableCreated) return;
			var command = _connection.CreateCommand();
			command.CommandText = $@"CREATE TABLE IF NOT EXISTS {_invRecordTable}
									({_idCol} INTEGER PRIMARY KEY ASC,
									{_qtyCol} INTEGER,
									{_itemCol} INTEGER,
									{_jobCol} INTEGER,
									{_dateCol} DATETIME);";
			command.ExecuteNonQuery();
			isTableCreated = true;
		}

	}

}
