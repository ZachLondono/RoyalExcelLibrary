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
using Microsoft.VisualBasic;

namespace RoyalExcelLibrary {
	public class ExcelLibrary {

#if DEBUG
            public const string db_path = "R:\\DB ORDERS\\RoyalExcelLibrary\\JobsTesting.db";
#else
            public const string db_path = "R:\\DB ORDERS\\RoyalExcelLibrary\\Jobs.db";
#endif

        public static void DrawerBoxProcessor(string format, bool trackjob, bool generateCutLists, bool printLabels, bool printCutlists, bool generatePackingList, bool printPackingList, bool generateInvoice, bool printInvoice, bool emailInvoice) {

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
                    string input = Interaction.InputBox("Enter Richelieu web number of order to process", "Web Number", "", 0, 0);
                    if (input.Equals("")) return;
                    provider = new RichelieuExcelDBOrderProvider(input);
                    googleExporter = new RichelieuGoogleSheetExport();
                    break;
                case "allmoxy":
                    filepath = ChooseFile();
                    if (filepath is null) return;
                    provider = new AllmoxyOrderProvider(filepath);

                    DialogResult result = MessageBox.Show("Is this an OT customer", "OT Customer", MessageBoxButtons.YesNo);

                    if (result == DialogResult.Yes) googleExporter = new OTGoogleSheetExport();
                    else googleExporter = new MetroGoogleSheetExport();
                    break;
                default:
                    throw new ArgumentException("Unknown provider format");
            }

            Order order;
            try {
                app.ScreenUpdating = false;
                order = provider.LoadCurrentOrder();
                app.ScreenUpdating = true;
            } catch (Exception e) {
                app.ScreenUpdating = true;
                errMessage.SetError("Error Loading Order", e.Message, e.ToString());
                errMessage.ShowDialog();
                return;
            }


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
                        var cutlists = productService.GenerateCutList(order, app.ActiveWorkbook, errMessage);

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

                        if (order.Job.JobSource.ToLower().Equals("hafele")) {
                            BOLExport bolExpt = new BOLExport();
                            var bol = bolExpt.ExportOrder(order, null, app.ActiveWorkbook);

                            if (printCutlists) bol.PrintOut(ActivePrinter: printerName);
                            else bol.PrintPreview();
                        }

                        app.ScreenUpdating = true;
                    } catch (Exception e) {
                        app.ScreenUpdating = true;

                        errMessage.SetError("Error While Cut Listing", e.Message, e.ToString());
                        errMessage.ShowDialog();
                    }

                }

                if (generateInvoice || generatePackingList) {

                    if (generatePackingList) {

                        try {

                            app.ScreenUpdating = false;
                            Worksheet packingList = productService.GeneratePackingList(order, app.ActiveWorkbook, errMessage);
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
                            errMessage.ShowDialog();
                        }

                    }

                    if (generateInvoice) {

                        try {

                            InvoiceExport invoiceExp = new InvoiceExport();

                            app.ScreenUpdating = false;
                            Worksheet invoice = productService.GenerateInvoice(order, app.ActiveWorkbook, errMessage);
                            app.ScreenUpdating = true;

                            if (printInvoice) {
                                if (!printerInstalled)
                                    // TODO open popup for user to select printer
                                    throw new InvalidOperationException($"Unable to print.\nPrinter '{printerName}' not available");
                                invoice.PrintOut(ActivePrinter: printerName);
                            } else invoice.PrintPreview();

                            if (emailInvoice) {
                                EmailArgs args = new EmailArgs {
                                    Subject = $"{order.Number} - Invoice",
                                    Body= "Please see attached invoice.",
                                    Attachments = new object[] { new AttachmentArgs { Source = invoice, DisplayName = "Invoice", FileName = $"{order.Number} - Invoice" } },
                                    AutoSend = false
                                };

#if DEBUG
                                args.From = "zach@royalcabinet.com";
#else
                                args.From = "dovetail@royalcabinet.com";
#endif

                                switch (order.Job.JobSource.ToLower()) {
                                    case "hafele":
                                        args.To = new string[] { "Accountspayable@hafele.us" };
                                        args.CC = new string[] { "Accounting@royalcabinet.com" };
                                        break;
                                    case "richelieu":
                                        args.To = new string[] { "AP@richelieu.com" };
                                        args.CC = new string[] { "Accounting@royalcabinet.com" };
                                        break;
                                    case "allmoxy":
                                        args.To = new string[] {"Accounting@royalcabinet.com"} ;
                                        break;
                                }

                                OutlookEmailExport.SendEmail(args);
                            }

                        } catch (Exception e) {
                            app.ScreenUpdating = true;
                            errMessage.SetError("Error While Generating/Printing Invoice", e.Message, e.ToString());
                            errMessage.ShowDialog();
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
                            break;
                        case "richelieu":
                            labelExport = new RichelieuLabelExport();
                            break;
                        case "ot":
                        default:
                            labelExport = new OTLabelExport();
                            break;
                    }

                    labelExport.PrintLables(order);

                } catch (Exception e) {
                    errMessage.SetError("Error While Printing Labels", e.Message, e.ToString());
                    errMessage.ShowDialog();
                }

			}

            initialWorksheet.Select();

            if (order.Job.JobSource.ToLower().Equals("allmoxy") || order.Job.JobSource.ToLower().Equals("hafele")) {
                try {
                    Range sourceRng = app.Range["OrderSource"];
                    if (order.Job.JobSource.ToLower().Equals("allmoxy")) {
                        sourceRng.Value2 = $"https://metrodrawerboxes.allmoxy.com/orders/quote/{order.Number}/";
                    } else if (order.Job.JobSource.ToLower().Equals("hafele")) {
                        var parts = filepath.Split('\\');
                        var filename = parts[parts.Length - 1];
                        sourceRng.Value2 = $"=HYPERLINK(\"{filepath}\", \"Open Source File [{filename}]\")";
                    }
                } catch (Exception e) {
                    errMessage.SetError("Error While Setting Job Source Link", e.Message, e.ToString());
                    errMessage.ShowDialog();
                }
            }

            errMessage.Dispose();

        }
        
		private static string ChooseFile() {
            var fileDialog = new OpenFileDialog();
            var result = fileDialog.ShowDialog();
            if (result != DialogResult.OK) return null;
            return fileDialog.FileName;
        }


        /// <summary>
        /// Calculates the stripe transaction fee, based on 2.45% + $0.30 processing fee and a 0.5% application fee
        /// </summary>
        /// <param name="totalCharge">The total transaction amount</param>
        /// <returns>The total transaction fee</returns>
        public static decimal CalculateStripeFee(decimal totalCharge) {
            // Stripe Fee = (total * 2.45%) + (total * 0.5%) + $0.30
            // multiply total in cents by percentage * 100 then divide by 10^4 to return to cents
            decimal processingFee = Math.Round(totalCharge * 0.0245M,2);
            decimal applicationFee = Math.Round(totalCharge * 0.0050M,2);
            decimal surcharge = 0.3M;

            return processingFee + applicationFee + surcharge;

        }

        /// <summary>
        /// Calculates the total commission for a transaction. Commission is calculated after deducting transaction fee, shipping fee and tax from the total transaction cost
        /// </summary>
        /// <param name="totalCharge">Total transaction ammount</param>
        /// <param name="shippingCost">Total shipping cost</param>
        /// <param name="tax">Total tax amount</param>
        /// <param name="commissionRate">Commission multiplier</param>
        /// <returns>The total commission to pay</returns>
        public static decimal CalculateCommissionPayment(decimal totalCharge, decimal shippingCost, decimal tax, decimal stripeFee, decimal commissionRate) {
            // Only earn commission on the net revenue, after fees, not including shipping or tax
            decimal commissionBase = totalCharge - stripeFee - shippingCost - tax;

            return Math.Round(commissionBase * commissionRate, 2, MidpointRounding.AwayFromZero);
		}

	}

}
