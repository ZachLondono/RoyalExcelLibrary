using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics;
using RoyalExcelLibrary.Providers;
using RoyalExcelLibrary.Services;
using RoyalExcelLibrary.Models;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Data.Sqlite;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using System.Data;
using RoyalExcelLibrary.DAL.Repositories;
using System.Windows.Forms;

namespace RoyalExcelLibrary {
	public class ExcelLibrary {

        private static readonly string db_path = "R:\\DB ORDERS\\RoyalExcelLibrary\\Jobs.db";

        public static async void DrawerBoxProcessor(string format) {

            var app = ExcelDnaUtil.Application as Excel.Application;

            Task<SqliteConnection> connTask = Task.Run(() => {
                return new SqliteConnection($"Data Source='{db_path}'");
            });

            Task<Order> orderTask = Task.Run(() => {

                IOrderProvider provider;
                switch (format.ToLower()) {
                    case "ot":
                        provider = new OTDBOrderProvider(app);
                        break;
                    case "hafele":
                        provider = new HafeleDBOrderProvider(app);
                        break;
                    default:
                        throw new ArgumentException("Unknown provider format");
                }

                return provider.LoadCurrentOrder();

            });

            SqliteConnection dbConnection = await connTask;
            Order order = await orderTask;

            using (dbConnection) {

                var productService = new DrawerBoxService(dbConnection);
                var inventoryService = new InventoryService(dbConnection);
                IEnumerable<Part> unplacedParts = null;

                dbConnection.Open();

                order = productService.StoreCurrentOrder(order);
                inventoryService.TrackMaterialUsage(order, out unplacedParts);
                
                dbConnection.Close();

                if (unplacedParts != null) {
                    string unplaced = "";
                    foreach (Part part in unplacedParts)
                        unplaced += $"{part.Qty}x{part.Width}Wx{part.Length}L {part.Material}\n";

                    if (!string.IsNullOrEmpty(unplaced))
                        MessageBox.Show("Unable to find available inventory for the following parts:\n" + unplaced, "Untracked Parts");
                }

            }

        }

    }
    
}
