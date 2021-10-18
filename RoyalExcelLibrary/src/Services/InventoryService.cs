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
		// Calculates the required material to complete the job and tracks that information 
		// </summary>
		public void TrackMaterialUsage(Order order, out IEnumerable<Part> unplacedParts) {

			double waste = 0.15;

			IEnumerable<InventoryUseRecord> itemsUsed = GetOptimizedParts(_inventoryRepository.GetAll(), order.Products, order.Job.Id, waste, out unplacedParts);
			foreach (InventoryUseRecord record in itemsUsed)
				_recordRepository.Insert(record);

		}

		public static IEnumerable<InventoryUseRecord> GetOptimizedParts(IEnumerable<InventoryItem> availableInventory, IEnumerable<IProduct> items, int jobId, double waste, out IEnumerable<Part> unplacedParts) {

			List<(InventoryItem, double)> offcuts = new List<(InventoryItem, double)>();
			List<InventoryUseRecord> itemsUsed = new List<InventoryUseRecord>();

			List<Part> unplaced = new List<Part>();

			DateTime timestamp = DateTime.Now;

			foreach (IProduct product in items) {
				foreach (Part part in product.GetParts().OrderByDescending((part) => part.Length)) {

					double materialUse = 0;

					Debug.WriteLine($"Cutting part: {part.Material}x{part.Qty}x{part.Width}x{part.Length}");

					if (part.UseType == InventoryUseType.Linear) {
						materialUse = part.Length * part.Qty;

						for (int p = 0; p < part.Qty; p++) {
							// First check if the piece can be cut from the available offcuts
							var availableOffcuts = offcuts.Where((offcut) => offcut.Item1.Material == part.Material && (offcut.Item1.Width - part.Width >= -1) && offcut.Item2 >= part.Length)
														.OrderByDescending((offcut) => offcut.Item2);

							
							Debug.WriteLine($"Available offcuts: {availableOffcuts.Count()}");

							var availableOffcut = availableOffcuts.FirstOrDefault();

							if (availableOffcut.Item1 != null) {

								Debug.WriteLine($"{availableOffcuts.First().Item1.Material} == {part.Material} ? {availableOffcuts.First().Item1.Material == part.Material}");

								for (int i = 0; i < offcuts.Count(); i++) {
									
									if (offcuts[i].Equals(availableOffcut)) {
										offcuts[i] = (availableOffcut.Item1, offcuts[i].Item2 - part.Length);
									}

								}
								
								Debug.WriteLine($"Using offcut '{availableOffcut.Item1.Name}'");

							} else {

								InventoryItem bestMatch = availableInventory.Where((item) => item.Material == part.Material && (item.Width - part.Width >= -1) && item.Length >= part.Length)
																	.OrderByDescending((item) => item.Length)
																	.OrderBy((item) => item.Width)
																	.FirstOrDefault();

								if (bestMatch is null) {
									Debug.WriteLine("Unable to find available inventory for part");
									unplaced.Add(part);
								} else {
									if (bestMatch.Length > part.Length) {
										double availableLength = bestMatch.Length - part.Length - (part.Length * waste);
										var offcut = (bestMatch, availableLength);
										offcuts.Add(offcut);
									}

									Debug.WriteLine($"Using additional material '{bestMatch.Name}'");
									var useRecord = itemsUsed.FirstOrDefault((item) => item.ItemId == bestMatch.Id);
									if (useRecord is null) {
										var newRecord = new InventoryUseRecord {
											ItemId = bestMatch.Id,
											Qty = 1,
											JobId = jobId,
											Timestamp = timestamp
										};
										itemsUsed.Add(newRecord);
									}  else useRecord.Qty++;

								}
							}
						}

					} else if (part.UseType == InventoryUseType.Area) {
						materialUse = part.Width * part.Length * part.Qty;
					}


				}
			}

			unplacedParts = unplaced;
			return itemsUsed;

		}

	}

}
