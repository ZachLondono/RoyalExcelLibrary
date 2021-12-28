using System;
using System.Collections.Generic;
using System.Linq;
using RoyalExcelLibrary.ExcelUI.Models;
using System.Data;
using System.Data.OleDb;
using RoyalExcelLibrary.Application.Features.Product;

namespace RoyalExcelLibrary.ExcelUI.src {

    public class Functions {

        public static object[] TestOrderQuery(int id) {

            var order = RoyalAddIn.QueryOrder(id);

            if (order == null) return new object[] { -1 };

            return new string[] {
                $"Customer: {order.Customer}",
                $"Name: {order.OrderName}",
                $"Numer: {order.OrderNumber}",
                $"Date: {order.OrderDate.ToString()}",
                $"Products: {order.Products.Count()}",
                $"Details: {order.OrderDetails.Count()}",
            };

        }

        public static int TestOrderStore() {

            var order = new Application.Features.Order.Order();
            order.Customer = "CustomerABC";
            order.OrderName = "OrderABC";
            order.OrderNumber = "123ABC";
            order.OrderDate = DateTime.Now;
            order.OrderDetails.Add("123", "ABC");

            var boxType = new Application.Features.Options.Materials.MaterialType();
            boxType.TypeId = 1;
            var bottType = new Application.Features.Options.Materials.MaterialType();
            bottType.TypeId = 2;
            List<IProduct> products = new List<IProduct>() {
                new DrawerBoxBuilder()
                    .WithQty(1)
                    .WithHeight(4.125)
                    .WithWidth(21)
                    .WithDepth(21)
                    .WithBoxMaterial(boxType)
                    .WithBotMaterial(bottType)
                    .WithExtra("UndermountNotch", "Standard Notch")
                    .WithExtra("Clips", "Blum")
                    .Build()
            };

            order.Products = products;

            order = RoyalAddIn.StoreOrder(order);

            return order.Id;

        }

        public static int TestDBStore() {

            var boxType = new Application.Features.Options.Materials.MaterialType();
            boxType.TypeId = 1;
            var bottType = new Application.Features.Options.Materials.MaterialType();
            bottType.TypeId = 2;

            DrawerBox box = new DrawerBoxBuilder()
                .WithQty(1)
                .WithHeight(4.125)
                .WithWidth(21)
                .WithDepth(21)
                .WithBoxMaterial(boxType)
                .WithBotMaterial(bottType)
                .WithExtra("UndermountNotch", "Standard Notch")
                .WithExtra("Clips", "Blum")
                .Build();

            box = RoyalAddIn.StoreDrawerBox(box, 123);
            return box.Id;

        }

        public static string TestDBQuery(int boxId) {

            DrawerBox box = RoyalAddIn.QueryDrawerBox(boxId);

            return $"{box.Height}H x {box.Width}W x {box.Depth}D : box={box.BoxMaterial.MaterialName}, bottom={box.BottomMaterial.MaterialName}";

        }

        public static int GetMaterials() {

            var mats = RoyalAddIn.GetMaterials();

            if (mats is null) return -1;

            return mats.Count();

        }

        public static object GetOptimizedParts(DateTime startDate, DateTime endDate) {

            // Add one day to the end date so that it includes orders on that date
            endDate = endDate.AddDays(1);

            IEnumerable<InventoryUseRecord> records;
            IEnumerable<InventoryItem> availableInventory; 
            using (var dbConnection = new OleDbConnection(ExcelLibrary.ConnectionString)) {
                dbConnection.Open();
                records = GetAllParts(dbConnection);
                availableInventory = GetAllAvailableInventory(dbConnection);
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

        private static IEnumerable<InventoryUseRecord> GetAllParts(OleDbConnection dbConnection) {

            using (OleDbCommand command = new OleDbCommand()) {

                command.CommandType = CommandType.Text;
                command.Connection = dbConnection;

                command.CommandText = $@"SELECT [Qty], [Material], [Width], [Length], [Thickness], [JobId], [Timestamp], [Id]
										FROM Parts;";

                List<InventoryUseRecord> records = new List<InventoryUseRecord>();

                using (var reader = command.ExecuteReader()) {

                    while (reader.Read()) {

                        string name = reader.GetString(1);
                        var e = Enum.Parse(typeof(MaterialType), name);

                        var itemRecord = new InventoryUseRecord {
                            Qty = reader.GetInt32(0),
                            Material = (MaterialType)Enum.Parse(typeof(MaterialType), reader.GetString(1)),
                            Width = reader.GetDouble(2),
                            Length = reader.GetDouble(3),
                            Thickness = reader.GetDouble(4),
                            JobId = reader.GetInt32(5),
                            Timestamp = reader.GetDateTime(6),
                            Id = reader.GetInt32(7)
                        };

                        records.Add(itemRecord);

                    }

                }

                return records;

            }

        }

        private static IEnumerable<InventoryItem> GetAllAvailableInventory(OleDbConnection dbConnection) {

            using (OleDbCommand command = new OleDbCommand()) {

                command.CommandType = CommandType.Text;
                command.Connection = dbConnection;

                command.CommandText = $@"SELECT [InventoryName], [Length], [Width], [Thickness], [Available], [Material], [Id]
									FROM [AvailableMaterial];";

                List<InventoryItem> items = new List<InventoryItem>();

                using (var reader = command.ExecuteReader()) {

                    while (reader.Read()) {

                        var item = new InventoryItem {
                            Name = reader.GetString(0),
                            Length = reader.GetDouble(1),
                            Width = reader.GetDouble(2),
                            Thickness = reader.GetDouble(3),
                            IsAvailable = reader.GetBoolean(4),
                            Material = (MaterialType)Enum.Parse(typeof(MaterialType), reader.GetString(5)),
                            Id = reader.GetInt32(6)
                        };

                        items.Add(item);

                    }

                }

                return items;
            }

        }


    }

}

