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

namespace ExcelLibTests {
	
	[TestClass]
	public class JobRepositoryTests {

		private readonly string _connString = "Data Source=InMemory;Mode=Memory;Cache=Shared";

		[TestMethod]
		public void AddJob() {

			using (SqliteConnection connection = new SqliteConnection(_connString)) {

				Job job = new Job();
				job.Id = 1;
				job.Name = "ABC123";
				job.CreationDate = DateTime.Now;

				connection.Open();
				JobRepository repo = new JobRepository(connection);
				repo.Insert(job);
				connection.Close();
			}

		}

	}
}
