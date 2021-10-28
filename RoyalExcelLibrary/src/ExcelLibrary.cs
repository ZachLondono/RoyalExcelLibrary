﻿using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using RoyalExcelLibrary.Providers;
using RoyalExcelLibrary.Services;
using RoyalExcelLibrary.Models;
using Microsoft.Data.Sqlite;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelDna.Integration;
using System.Windows.Forms;
using RoyalExcelLibrary.ExportFormat.Labels;
using RoyalExcelLibrary.ExportFormat;
using Microsoft.Office.Interop.Excel;

namespace RoyalExcelLibrary {
	public class ExcelLibrary {

        #if DEBUG
            public const string db_path = "R:\\DB ORDERS\\RoyalExcelLibrary\\JobsTesting.db";
#else
            public const string db_path = "R:\\DB ORDERS\\RoyalExcelLibrary\\Jobs.db";
#endif

        public static void DrawerBoxProcessor(string format, bool trackjob, bool generateCutLists, bool printLabels, bool printCutlists, bool generatePackingList, bool printPackingList) {

#if DEBUG
            MessageBox.Show($"Starting in Debug Mode\n Using database: '{db_path}'");
#endif

            var app = ExcelDnaUtil.Application as Excel.Application;

            IOrderProvider provider;
            switch (format.ToLower()) {
                case "ot":
                    provider = new OTDBOrderProvider(app);
                    break;
                case "hafele":
                    provider = new HafeleDBOrderProvider(app);
                    break;
                case "richelieu":
                    provider = new RichelieuExcelDBOrderSource(app);
                    break;
                case "allmoxy":
                    var fileDialog = new OpenFileDialog();
                    var result = fileDialog.ShowDialog();
                    if (result != DialogResult.OK) return;
                    string filepath = fileDialog.FileName;
                    provider = new AllmoxyOrderProvider(filepath);
                    break;
                default:
                    throw new ArgumentException("Unknown provider format");
            }

            Order order = provider.LoadCurrentOrder();

            bool printerInstalled = false;
            string printerName = "SHARP MX-M283N PCL6";
            if (printCutlists || printPackingList) {
                var printers = System.Drawing.Printing.PrinterSettings.InstalledPrinters;

                foreach (var printer in printers) {
                    if (printer.Equals(printerName)) {
                        printerInstalled = true;
                        break;
                    }
                }
            }

            if (trackjob || generateCutLists) {
                SqliteConnection dbConnection = new SqliteConnection($"Data Source='{db_path}'");

                using (dbConnection) {

                    var productService = new DrawerBoxService(dbConnection);
                    var inventoryService = new InventoryService(dbConnection);
                    IEnumerable<Part> unplacedParts = null;

                    if (trackjob) {
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

                    if (generateCutLists) {

                        try {
                            app.ScreenUpdating = false;
                            IProductService service = new DrawerBoxService(dbConnection);
                            dbConnection.Open();
                            var cutlists = service.GenerateCutList(order, app.ActiveWorkbook);
                            dbConnection.Close();

                            if (!printerInstalled) {
                                // TODO open popup for user to select printer
                                throw new InvalidOperationException($"Unable to print.\nPrinter '{printerName}' not available");
                            }

                            foreach (var cutlist in cutlists) {
                                if (printCutlists)
                                    cutlist.PrintOut(ActivePrinter: printerName);    
                                else cutlist.PrintPreview();
                            }

                            app.ScreenUpdating = true;
                        } catch (Exception e) {
                            app.ScreenUpdating = true;
                            var result = MessageBox.Show($"An error occured while processing drawer boxes\nShow error message?\n[{e.Message}]", "Error occurred", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                            if (result == DialogResult.Yes) {
                                MessageBox.Show(e.ToString(), "Error Message");
                            }
                        }

                    }


                }
            }

            if (generatePackingList) {

                try {
                    PackingListExport packingListExp = new PackingListExport();

                    PackingListData data = new PackingListData {

                        SupplierName = "Metro Drawer Boxes",
                        SupplierContact = "",
                        SupplierAddress = new Address {
                            StreetAddress = "15E Easy St",
                            City = "Bound Brook",
                            State = "NJ",
                            Zip = "08805"
                        },

                        RecipientName = order.CustomerName,
                        RecipientContact = "",
                        RecipientAddress = order.ShipAddress

                    };

                    app.ScreenUpdating = false;
                    Worksheet packingList = packingListExp.ExportOrder(order, data, app.ActiveWorkbook);
                    app.ScreenUpdating = true;

                    if (!printerInstalled) {
                        // TODO open popup for user to select printer
                        throw new InvalidOperationException($"Unable to print.\nPrinter '{printerName}' not available");
                    }

                    if (printPackingList) packingList.PrintOut(ActivePrinter: printerName);    
                    else packingList.PrintPreview();
                } catch (Exception e) {
                    app.ScreenUpdating = true;
                    var result = MessageBox.Show($"An error occured while generating/printing the packing list\nShow error message?\n[{e.Message}]", "Error occurred", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                    if (result == DialogResult.Yes) {
                        MessageBox.Show(e.ToString(), "Error Message");
                    }

                }

            }

            if (printLabels) {

                try {

                    ILabelExport labelExport;
                    switch (format.ToLower()) {
                        case "hafele":
                            labelExport = new HafeleLabelExport();
                            (labelExport as HafeleLabelExport).ProjectNum = app.Range["Order!J7"].Value2.ToString();
                            break;
                        case "ot":
                        default:
                            labelExport = new OTLabelExport();
                            break;
                    }

                    labelExport.PrintLables(order);

                } catch (Exception e) {
                    var result = MessageBox.Show($"An error occured while printing labels\nShow error message?\n[{e.Message}]", "Error occurred", MessageBoxButtons.YesNo, MessageBoxIcon.Error);
                    if (result == DialogResult.Yes)
                        MessageBox.Show(e.ToString(), "Error Message");
                }

			}

        }
        
	}

}
