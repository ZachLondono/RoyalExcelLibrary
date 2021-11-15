using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using RoyalExcelLibrary.DAL;
using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;

using Microsoft.Extensions.Logging;
using System.Data;
using RoyalExcelLibrary.DAL.Repositories;

namespace RoyalExcelLibrary.Services {

	public class InventoryService {

		private readonly IDbConnection _connection;
		private readonly IInventoryRecordRepository _recordRepository;
		private readonly IInventoryRepository _inventoryRepository;

		public InventoryService(IDbConnection connection) {
			_connection = connection;
			_recordRepository = new InventoryRecordRepository(connection);
			_inventoryRepository = new InventoryRepository(connection);
		}

		// <summary>
		// Store all parts used in the job
		// </summary>
		public void TrackMaterialUsage(Order order) {

			DateTime trackTime = DateTime.Now;
			foreach (Product prod in order.Products) {
				foreach (Part part in prod.GetParts()) {

					var record = new InventoryUseRecord {
						Qty = part.Qty,
						Width = part.Width,
						Length = part.Length,
						Thickness = 0,
						JobId = order.Job.Id,
						Material = part.Material,
						Timestamp = trackTime
					};


					_recordRepository.Insert(record);
				}
			}

		}

	}

}
