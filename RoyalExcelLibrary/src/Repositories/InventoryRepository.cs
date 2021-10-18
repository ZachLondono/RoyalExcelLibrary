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
	public class InventoryRepository : IInventoryRepository {

		private readonly string _inventoryTable = "available_inventory";
		private readonly string _idCol = "id";
		private readonly string _nameCol = "name";
		private readonly string _widthCol = "width";
		private readonly string _lengthCol = "length";
		private readonly string _thicknessCol = "thickness";
		private readonly string _availableCol = "available";
		private readonly string _materialCol = "material_type";

		private bool isTableCreated = false;
		private readonly IDbConnection _connection;

		public InventoryRepository(IDbConnection connection) {
			_connection = connection;
		}

		public void Delete(InventoryItem entity) {

			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $@"DELETE FROM {_inventoryTable} 
									WHERE {_idCol} = @id;";

			command.AddParamWithValue("@id", entity.Id);

			try {
				if (command.ExecuteNonQuery() != 1) throw new InvalidOperationException("Unable to delete entity");
			} catch (Exception e) {
				Debug.WriteLine(e);
			}


		}

		public InventoryItem GetById(int id) {

			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $@"SELECT {_nameCol}, {_lengthCol}, {_widthCol}, {_thicknessCol}, {_availableCol}, {_materialCol}	
									FROM {_inventoryTable}
									WHERE {_idCol} = @id;";

			command.AddParamWithValue("@id", id);

			using (var reader = command.ExecuteReader()) {

				reader.Read();

				var item = new InventoryItem {
					Name = reader.GetString(0),
					Length = reader.GetDouble(1),
					Width = reader.GetDouble(2),
					Thickness = reader.GetDouble(3),
					IsAvailable = reader.GetBoolean(4),
					Material = MaterialFunctions.StringToType(reader.GetString(5)),
					Id = id
				};

				return item;

			}

			throw new InvalidOperationException($"Entity with id '{id}' does not exist");

		}

		public InventoryItem Insert(InventoryItem entity) {

			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $@"INSERT INTO {_inventoryTable} 
									({_nameCol}, {_lengthCol}, {_widthCol}, {_thicknessCol}, {_availableCol}, {_materialCol})
									VALUES
									(@name, @length, @width, @thickness, @availability, @material);
									SELECT last_insert_rowid();";
			command.AddParamWithValue("@name", entity.Name);
			command.AddParamWithValue("@length", entity.Length);
			command.AddParamWithValue("@width", entity.Width);
			command.AddParamWithValue("@thickness", entity.Thickness);
			command.AddParamWithValue("@availability", entity.IsAvailable);
			command.AddParamWithValue("@material", MaterialFunctions.TypeToString(entity.Material));

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

		public void Update(InventoryItem entity) {
			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $@"UPDATE {_inventoryTable}
									SET {_nameCol} = @name, {_widthCol} = @width, {_lengthCol} = @length, {_thicknessCol} = @thickness, {_availableCol} = @availability, {_materialCol} = @material
									WHERE {_idCol} = @id;";
			command.AddParamWithValue("@name", entity.Name);
			command.AddParamWithValue("@length", entity.Length);
			command.AddParamWithValue("@width", entity.Width);
			command.AddParamWithValue("@thickness", entity.Thickness);
			command.AddParamWithValue("@availability", entity.IsAvailable);
			command.AddParamWithValue("@material", MaterialFunctions.TypeToString(entity.Material));

			if (command.ExecuteNonQuery() != 1) throw new InvalidOperationException("Unable to update record");

		}

		public IEnumerable<InventoryItem> GetAll() {
			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $@"SELECT {_nameCol}, {_lengthCol}, {_widthCol}, {_thicknessCol}, {_availableCol}, {_materialCol}	, {_idCol}
									FROM {_inventoryTable};";

			List<InventoryItem> items = new List<InventoryItem>();

			using (var reader = command.ExecuteReader()) {

				while (reader.Read()) {

					var item = new InventoryItem {
						Name = reader.GetString(0),
						Length = reader.GetDouble(1),
						Width = reader.GetDouble(2),
						Thickness = reader.GetDouble(3),
						IsAvailable = reader.GetBoolean(4),
						Material = MaterialFunctions.StringToType(reader.GetString(5)),
						Id = reader.GetInt32(5)
					};

					items.Add(item);

				}

			}

			return items;

		}

		private void CreateTable() {
			if (isTableCreated) return;
			var command = _connection.CreateCommand();
			command.CommandText = $@"CREATE TABLE IF NOT EXISTS {_inventoryTable}
									({_idCol} INTEGER PRIMARY KEY ASC,
									{_nameCol} VARCHAR,
									{_widthCol} DOUBLE,
									{_lengthCol} DOUBLE,
									{_thicknessCol} DOUBLE,
									{_availableCol} BOOLEAN,
									{_materialCol} VARCHAR);";

			command.ExecuteNonQuery();
			isTableCreated = true;
		}

	}

}
