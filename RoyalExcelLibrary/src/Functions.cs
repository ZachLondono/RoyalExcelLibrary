using System;
using System.Collections.Generic;
using System.Linq;
using RoyalExcelLibrary.Models;
using Microsoft.Data.Sqlite;
using System.Data;
using RoyalExcelLibrary.DAL.Repositories;
using RoyalExcelLibrary.Models.Products;

namespace RoyalExcelLibrary.src {
    public class Functions {

        public static object GetOptimizedParts(DateTime startDate, DateTime endDate) {

            // Add one day to the end date so that it includes orders on that date
            endDate = endDate.AddDays(1);

            IEnumerable<InventoryUseRecord> records;
            IEnumerable<InventoryItem> availableInventory; 
            using (SqliteConnection dbConnection = new SqliteConnection($"Data Source='{ExcelLibrary.db_path}'")) {
                dbConnection.Open();
                records = new InventoryRecordRepository(dbConnection).GetAll();
                availableInventory = new InventoryRepository(dbConnection).GetAll();
            }

            // Filter records by date
            records = records.Where(r => r.Timestamp >= startDate && r.Timestamp <= endDate);

            // Optimize the material, given the available inventory
            IList<MatUse> matUse = OptimizeFromInventory(records, availableInventory);

            matUse = matUse.OrderByDescending(m => m.Width)
                .OrderByDescending(m => m.Length)
                .OrderBy(m => m.Material)
                .ToList();

            var table = new string[matUse.Count + 1, typeof(MatUse).GetProperties().Count()];
            table[0, 0] = "Material Type";
            table[0, 1] = "Qty";
            table[0, 2] = "Width / Height";
            table[0, 3] = "Length";

            for (int i = 0; i < matUse.Count; i++) {
                table[i + 1, 0] = matUse[i].Material;
                table[i + 1, 1] = matUse[i].Qty.ToString();
                table[i + 1, 2] = matUse[i].Width.ToString();
                table[i + 1, 3] = matUse[i].Length.ToString();
            }

            return table;

        }

        private static IList<MatUse> OptimizeFromInventory(IEnumerable<InventoryUseRecord> records, IEnumerable<InventoryItem> inventoryItems) {

            // Start from the largest part to the smallest part
            var orderedRecords = records.OrderByDescending(r => r.Length);
            // Use the smallest available inventory first
            var orderedInventory = inventoryItems.OrderBy(i => i.Width)
                                                 .OrderBy(i => i.Length);

            Dictionary<InventoryItem, int> itemUse = new Dictionary<InventoryItem, int>();
            HashSet<double> stdHeights = new HashSet<double> {

            };

            var recordsByJob = records.GroupBy(r => r.JobId);

            foreach (var jobrecords in recordsByJob) {
                List<(InventoryItem, double)> offcuts = new List<(InventoryItem, double)>();
            
                foreach (var record in jobrecords) {

                    bool offcutUsed = false;
                    // Look for an existing offcut to use
                    foreach (var offcut in offcuts) {

                        InventoryItem offcutItem = offcut.Item1;
                        double length = offcut.Item2;

                        if (offcutItem.Material == record.Material &&
                            ((offcutItem.Width - record.Width == 0.5) || (!stdHeights.Contains(record.Width) && offcutItem.Width > record.Width))) {

                            var leftover = length - record.Length;
                            offcuts.Remove(offcut);
                            if (leftover > 100)
                                offcuts.Add((offcutItem, leftover));

                            offcutUsed = true;
                            break;
                        }
                    }

                    // If an offcut is used, there is no need to get another piece of material from inventory
                    if (offcutUsed) continue;

                    foreach (var item in orderedInventory) {

                        if (!item.IsAvailable || item.Material != record.Material) continue;

                        if (item.Length >= record.Length && ((item.Width - record.Width == 0.5) || (!stdHeights.Contains(record.Width) && item.Width > record.Width))) {

                            // Add one to the total quantity
                            int qty = 0;
                            if (itemUse.ContainsKey(item))
                                qty = itemUse[item];
                            else itemUse.Add(item, 0);
                            itemUse[item] = qty + 1;

                            // Add extra to offcuts
                            if (item.Length > record.Length && item.Length - record.Length > 100) {
                                offcuts.Add((item, item.Length - record.Length));
                            }

                            break;
                        }

                    }
                }
            }

            IList<MatUse> matUse = new List<MatUse>();

            foreach (var use in itemUse) {
                matUse.Add(new MatUse {
                    Material = Enum.GetName(typeof(MaterialType), use.Key.Material),
                    Length = use.Key.Length,
                    Width = use.Key.Width,
                    Qty = use.Value
                });
            }

            return matUse;

        }

        private struct MatUse {

            public string Material { get; set; }

            public double Width { get; set; }

            public double Length { get; set; }

            public int Qty { get; set; }

        }

    }

}

