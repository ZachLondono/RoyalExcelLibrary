using System;
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
using System.Diagnostics;
using RoyalExcelLibrary.Views;
using RoyalExcelLibrary.ExportFormat.Google;

namespace RoyalExcelLibrary {
	public class ExcelLibrary {

#if DEBUG
            public const string db_path = "R:\\DB ORDERS\\RoyalExcelLibrary\\JobsTesting.db";
#else
            public const string db_path = "R:\\DB ORDERS\\RoyalExcelLibrary\\Jobs.db";
#endif

        public static void DrawerBoxProcessor(string format, bool trackjob, bool generateCutLists, bool printLabels, bool printCutlists, bool generatePackingList, bool printPackingList, bool generateInvoice, bool printInvoice) {

#if DEBUG
            MessageBox.Show($"Starting in Debug Mode\n Using database: '{db_path}'");
#endif

            ErrorMessage errMessage = new ErrorMessage();
            errMessage.TopMost = true;

            Excel.Application app = ExcelDnaUtil.Application as Excel.Application;
            Worksheet initialWorksheet = app.ActiveSheet;

            IOrderProvider provider;
            IGoogleSheetsExport googleExporter;
            string filepath = null;
            switch (format.ToLower()) {
                case "ot":
                    provider = new OTDBOrderProvider(app);
                    googleExporter = new OTGoogleSheetExport();
                    break;
                case "hafele":
                    filepath = ChooseFile();
                    if (filepath is null) return;
                    provider = new HafeleDBOrderProvider(filepath);
                    googleExporter = new HafeleGoogleSheetExport();
                    break;
                case "richelieu":
                    provider = new RichelieuExcelDBOrderSource(app);
                    googleExporter = new RichelieuGoogleSheetExport();
                    break;
                case "allmoxy":
                    filepath = ChooseFile();
                    if (filepath is null) return;
                    provider = new AllmoxyOrderProvider(filepath);
                    googleExporter = new OTGoogleSheetExport();
                    break;
                default:
                    throw new ArgumentException("Unknown provider format");
            }

            app.ScreenUpdating = false;
            Order order = provider.LoadCurrentOrder();
            app.ScreenUpdating = true;

            // Check if the printer is available to print from
            bool printerInstalled = false;
            string printerName = "SHARP MX-M283N PCL6";
            var printers = System.Drawing.Printing.PrinterSettings.InstalledPrinters;

            foreach (var printer in printers) {
                if (printer.Equals(printerName)) {
                    printerInstalled = true;
                    break;
                }
            }

            SqliteConnection dbConnection = new SqliteConnection($"Data Source='{db_path}'");

            using (dbConnection) {

                IProductService productService = new DrawerBoxService(dbConnection);
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

                    googleExporter.ExportOrder(order);
                }

                if (generateCutLists) {

                    try {
                        app.ScreenUpdating = false;
                        var cutlists = productService.GenerateCutList(order, app.ActiveWorkbook);

                        if (trackjob) {
                            dbConnection.Open();
                            productService.SetOrderStatus(order, Status.Released);
                            dbConnection.Close();
                        }

                        if (!printerInstalled && printCutlists) {
                            // TODO open popup for user to select printer
                            throw new InvalidOperationException($"Unable to print.\nPrinter '{printerName}' not available");
                        }

                        foreach (var cutlist in cutlists) {
                            if (cutlist is null) continue;
                            if (printCutlists)
                                cutlist.PrintOut(ActivePrinter: printerName);    
                            else cutlist.PrintPreview();
                        }

                        app.ScreenUpdating = true;
                    } catch (Exception e) {
                        app.ScreenUpdating = true;

                        errMessage.SetError("Error While Cut Listing", e.Message, e.ToString());
                        errMessage.Show();
                    }

                }

                if (generateInvoice || generatePackingList) {

                    if (generatePackingList) {

                        try {

                            app.ScreenUpdating = false;
                            Worksheet packingList = productService.GeneratePackingList(order, app.ActiveWorkbook);
                            app.ScreenUpdating = true;

                            if (!printerInstalled && printPackingList) {
                                // TODO open popup for user to select printer
                                throw new InvalidOperationException($"Unable to print.\nPrinter '{printerName}' not available");
                            }

                            if (printPackingList) packingList.PrintOut(ActivePrinter: printerName);
                            else packingList.PrintPreview();

                        } catch (Exception e) {
                            app.ScreenUpdating = true;
                            errMessage.SetError("Error While Generating/Printing Packing List", e.Message, e.ToString());
                            errMessage.Show();
                        }

                    }

                    if (generateInvoice) {

                        try {

                            InvoiceExport invoiceExp = new InvoiceExport();

                            app.ScreenUpdating = false;
                            Worksheet invoice = productService.GenerateInvoice(order, app.ActiveWorkbook);
                            app.ScreenUpdating = true;

                            if (!printerInstalled) {
                                // TODO open popup for user to select printer
                                throw new InvalidOperationException($"Unable to print.\nPrinter '{printerName}' not available");
                            }

                            if (printPackingList) invoice.PrintOut(ActivePrinter: printerName);
                            else invoice.PrintPreview();

                        } catch (Exception e) {
                            app.ScreenUpdating = true;
                            errMessage.SetError("Error While Generating/Printing Invoice", e.Message, e.ToString());
                            errMessage.Show();
                        }

                    }
                }

            }

            if (printLabels) {

                try {

                    ILabelExport labelExport;
                    switch (format.ToLower()) {
                        case "hafele":
                            labelExport = new HafeleLabelExport();
                            //(labelExport as HafeleLabelExport).ProjectNum = app.Range["Order!J7"].Value2.ToString();
                            break;
                        case "ot":
                        default:
                            labelExport = new OTLabelExport();
                            break;
                    }

                    labelExport.PrintLables(order);

                } catch (Exception e) {
                    errMessage.SetError("Error While Printing Labels", e.Message, e.ToString());
                    errMessage.Show();
                }

			}

            initialWorksheet.Select();

            try {
                Range sourceRng = app.Range["OrderSource"];
                if (order.Job.JobSource.ToLower().Equals("allmoxy")) {
                    sourceRng.Value2 = $"https://metrodrawerboxes.allmoxy.com/orders/quote/{order.Number}/";
                } else if (order.Job.JobSource.ToLower().Equals("hafele")) {
                    sourceRng.Formula = $"=HYPERLINK('{filepath}', \"Open Source File\")";
                }
            } catch (Exception e) {
                errMessage.SetError("Error While Setting Job Source Link", e.Message, e.ToString());
                errMessage.Show();
            }

            errMessage.Dispose();

        }

        private static string ChooseFile() {
            var fileDialog = new OpenFileDialog();
            var result = fileDialog.ShowDialog();
            if (result != DialogResult.OK) return null;
            return fileDialog.FileName;
        }
        
	}

}
