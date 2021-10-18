using RoyalExcelLibrary.Extensions;
using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.DAL.Repositories {

	public class DrawerBoxRepository : IDrawerBoxRepository {

		private readonly string _dbTableName = "drawer_boxes";
		private readonly string _heightCol = "height";
		private readonly string _widthCol = "width";
		private readonly string _depthCol = "depth";
		private readonly string _qtyCol = "qty";
		private readonly string _sideCol = "side_material";
		private readonly string _bottCol = "bottom_material";
		private readonly string _idCol = "id";
		private readonly string _jobCol = "job_id";


		private readonly IDbConnection _connection;
		private bool isTableCreated;

		public DrawerBoxRepository(IDbConnection connection) {
			_connection = connection;
			isTableCreated = false;
		}

		public void Delete(DrawerBox entity) {

			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $@"DELETE FROM {_dbTableName} 
									WHERE {_idCol} = @id;";

			command.AddParamWithValue("@id", entity.Id);

			try {
				if (command.ExecuteNonQuery() != 1) throw new InvalidOperationException("Unable to delete entity");
			} catch (Exception e) {
				Debug.WriteLine(e);
			}

		}

		public DrawerBox GetById(int id) {

			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $@"SELECT {_qtyCol}, {_heightCol}, {_widthCol}, {_depthCol}, {_sideCol}, {_bottCol}, {_jobCol}	
									FROM {_dbTableName}
									WHERE {_idCol} = @id;";

			command.AddParamWithValue("@id", id);

			using (var reader = command.ExecuteReader()) {

				reader.Read();

				var drawerbox = new DrawerBox {
					Qty = reader.GetInt32(0),
					Height = reader.GetDouble(1),
					Width = reader.GetDouble(2),
					Depth = reader.GetDouble(3),
					SideMaterial = MaterialFunctions.StringToType(reader.GetString(4)),
					BottomMaterial = MaterialFunctions.StringToType(reader.GetString(5)),
					JobId = reader.GetInt32(6),
					Id = id
				};

				return drawerbox;

			}

			throw new InvalidOperationException($"Entity with id '{id}' does not exist");

		}

		public DrawerBox Insert(DrawerBox entity) {

			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $@"INSERT INTO 
										{_dbTableName} 
										({_qtyCol}, {_heightCol}, {_widthCol}, {_depthCol}, {_sideCol}, {_bottCol}, {_jobCol}) 
									VALUES
										(@qty, @height, @width, @depth, @side, @bott, @job);
									SELECT last_insert_rowid();";

			command.AddParamWithValue("@qty", entity.Qty);
			command.AddParamWithValue("@height", entity.Height);
			command.AddParamWithValue("@width", entity.Width);
			command.AddParamWithValue("@depth", entity.Depth);
			command.AddParamWithValue("@side", MaterialFunctions.TypeToString(entity.SideMaterial));
			command.AddParamWithValue("@bott", MaterialFunctions.TypeToString(entity.BottomMaterial));
			command.AddParamWithValue("@job", entity.JobId);
			
			using (var reader = command.ExecuteReader()) {

				if (!reader.Read()) throw new InvalidOperationException("Unable to insert entity");

				entity.Id = reader.GetInt32(0);

			}

			return entity;

		}

		public void Update(DrawerBox entity) {

			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $@"UPDATE {_dbTableName}
									SET {_qtyCol} = @qty, {_heightCol} = @height, {_widthCol} = @width, {_depthCol} = @depth, {_sideCol} = @side, {_bottCol} = @bott, {_jobCol} = @job
									WHERE {_idCol} = @id;";

			command.AddParamWithValue("@id", entity.Id);
			command.AddParamWithValue("@qty", entity.Qty);
			command.AddParamWithValue("@height", entity.Height);
			command.AddParamWithValue("@width", entity.Width);
			command.AddParamWithValue("@depth", entity.Depth);
			command.AddParamWithValue("@side", MaterialFunctions.TypeToString(entity.SideMaterial));
			command.AddParamWithValue("@bott", MaterialFunctions.TypeToString(entity.BottomMaterial));
			command.AddParamWithValue("@job", entity.JobId);

			if (command.ExecuteNonQuery() != 1) throw new InvalidOperationException($"Unable to update entity with id '{entity.Id}'");

		}

		public IEnumerable<DrawerBox> GetAll() {

			var command = _connection.CreateCommand();
			command.CommandText = $@"SELECT {_qtyCol}, {_heightCol}, {_widthCol}, {_depthCol}, {_sideCol}, {_bottCol}, {_jobCol}, {_idCol}
									FROM {_dbTableName};";

			List<DrawerBox> items = new List<DrawerBox>();

			using (var reader = command.ExecuteReader()) {

				while (reader.Read()) {

					var item = new DrawerBox {
						Qty = reader.GetInt32(0),
						Height = reader.GetDouble(1),
						Width = reader.GetDouble(2),
						Depth = reader.GetDouble(3),
						SideMaterial = MaterialFunctions.StringToType(reader.GetString(4)),
						BottomMaterial = MaterialFunctions.StringToType(reader.GetString(5)),
						JobId = reader.GetInt32(6),
						Id = reader.GetInt32(7)
					};

					items.Add(item);

				}

			}

			return items;

		}

		private void CreateTable() {
			if (isTableCreated) return;
			var command = _connection.CreateCommand();
			command.CommandText = $@"CREATE TABLE IF NOT EXISTS {_dbTableName}
									({_idCol} INTEGER PRIMARY KEY ASC,
									{_qtyCol} INTEGER,
									{_heightCol} DOUBLE,
									{_widthCol} DOUBLE,
									{_depthCol} DOUBLE,
									{_sideCol} VARCHAR,
									{_bottCol} VARCHAR,
									{_jobCol} INTEGER);";
			command.ExecuteNonQuery();
			isTableCreated = true;
		}

	}


}
