using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Extensions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.DAL.Repositories {
	public class JobRepository : IJobRepository {

		private readonly string _jobTableName = "jobs";
		private readonly string _jobIdCol = "id";
		private readonly string _jobNameCol = "name";
		private readonly string _jobDateCol = "date_created";
		private readonly string _jobRevenueCol = "gross_revenue";

		private readonly IDbConnection _connection;
		private bool isTableCreated;

		public JobRepository(IDbConnection connection) {
			_connection = connection;
			isTableCreated = false;
		}

		public void Delete(Job entity) {

			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $"DELETE FROM {_jobTableName} WHERE {_jobIdCol} = @id;";
			command.AddParamWithValue("@id", entity.Id);

			if (command.ExecuteNonQuery() != 1) throw new InvalidOperationException("Entity was not removed");

		}

		public Job GetById(int id) {

			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $"SELECT {_jobNameCol}, {_jobRevenueCol}, {_jobDateCol} FROM {_jobTableName} WHERE {_jobIdCol} = @id;";
			command.AddParamWithValue("@id", id);

			Job job;
			using (var reader = command.ExecuteReader()) {

				if (!reader.Read()) throw new InvalidOperationException($"No record exists with the id {id}");

				job = new Job {
					Name = reader.GetString(0),
					GrossRevenue = reader.GetDouble(1),
					CreationDate = reader.GetDateTime(2),
					Id = id
				};

			}

			return job;

		}

		public Job Insert(Job entity) {

			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $"INSERT INTO {_jobTableName} ({_jobNameCol}, {_jobRevenueCol}, {_jobDateCol}) VALUES (@name, @revenue, @date); SELECT last_insert_rowid();";
			command.AddParamWithValue("@name", entity.Name);
			command.AddParamWithValue("@revenue", entity.GrossRevenue);
			command.AddParamWithValue("@date", entity.CreationDate);

			using (var reader = command.ExecuteReader()) {

				if(!reader.Read()) {
					Debug.WriteLine("noting returned");
					return null;
				}
				
				int newId = reader.GetInt32(0);

				entity.Id = newId;

			}

			return entity;

		}

		public void Update(Job entity) {

			CreateTable();

			var command = _connection.CreateCommand();
			command.CommandText = $"UPDATE {_jobTableName} SET {_jobNameCol} = @name, {_jobRevenueCol} = @revenue, {_jobDateCol} = @date WHERE {_jobIdCol} = @id;";
			command.AddParamWithValue("@id", entity.Id);
			command.AddParamWithValue("@name", entity.Name);
			command.AddParamWithValue("@revenue", entity.GrossRevenue);
			command.AddParamWithValue("@date", entity.CreationDate);

			if (command.ExecuteNonQuery() != 1) throw new InvalidOperationException("Unable to update job");

		}

		public IEnumerable<Job> GetAll() {

			var command = _connection.CreateCommand();
			command.CommandText = $"SELECT {_jobIdCol}, {_jobNameCol}, {_jobDateCol} FROM {_jobTableName};";

			List<Job> jobs = new List<Job>();
			using (var reader = command.ExecuteReader()) {

				while (reader.Read()) {

					var job = new Job() {
						Id = reader.GetInt32(0),
						Name = reader.GetString(1),
						GrossRevenue = reader.GetDouble(2),
						CreationDate = reader.GetDateTime(3)
					};

					jobs.Add(job);

				}

			}


			return jobs;

		}

		private void CreateTable() {
			if (isTableCreated) return;
			var command = _connection.CreateCommand();
			command.CommandText = $@"CREATE TABLE IF NOT EXISTS {_jobTableName} 
										({_jobIdCol} INTEGER PRIMARY KEY ASC,
										{_jobNameCol} VARCHAR,
										{_jobRevenueCol} DOUBLE,
										{_jobDateCol} DATETIME);";
			command.ExecuteNonQuery();
			isTableCreated = true;
		}

	}

}
