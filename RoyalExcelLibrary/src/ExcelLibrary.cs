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
            string filepath = null;
            switch (format.ToLower()) {
                case "ot":
                    provider = new OTDBOrderProvider(app);
                    break;
                case "hafele":
                    filepath = ChooseFile();
                    if (filepath is null) return;
                    provider = new HafeleDBOrderProvider(filepath);
                    break;
                case "richelieu":
                    string input = Interaction.InputBox("Enter Richelieu web number of order to process", "Web Number", "", 0, 0);
                    if (input.Equals("")) return;
                    provider = new RichelieuExcelDBOrderProvider(input);
                    break;
                case "allmoxy":
                    filepath = ChooseFile();
                    if (filepath is null) return;
                    provider = new AllmoxyOrderProvider(filepath);
                    break;
                case "loaded":
                    provider = new UniversalDBOrderProvider(app);
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

                if (trackjob) {
                    dbConnection.Open();
                    order = productService.StoreCurrentOrder(order);
                    inventoryService.TrackMaterialUsage(order);
                    dbConnection.Close();

                    switch (order.Job.JobSource) {
                        case "hafele":
                            new HafeleGoogleSheetExport().ExportOrder(order);
                            break;
                        case "richlieu":
                            new RichelieuGoogleSheetExport().ExportOrder(order);
                            break;
                        case "ot":
                            new OTGoogleSheetExport().ExportOrder(order);
                            break;
                        case "allmoxy":
                            DialogResult result = MessageBox.Show("Is this an OT customer", "OT Customer", MessageBoxButtons.YesNo);
                            if (result == DialogResult.Yes) new OTGoogleSheetExport().ExportOrder(order);
                            else new MetroGoogleSheetExport().ExportOrder(order);
                            break;
                    }
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
                            var bol = bolExpt.ExportOrder(order, app.ActiveWorkbook);

                            if (printCutlists) {
                                bol.PrintOut(ActivePrinter: printerName);

                                DymoLabelService labelService = new DymoLabelService("R:\\DB ORDERS\\Labels\\Duie Pyle notice.label");
                                labelService.AddLabel(labelService.CreateLabel(), 1);
                                labelService.PrintLabels();

                            } else bol.PrintPreview();
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

            errMessage.Dispose();

        }

        /// <summary>
        /// Loads an order from provider and writes it to the excel worksheet
        /// </summary>
        /// <param name="providerName"></param>
        public static void LoadOrder(string providerName) {

            ErrorMessage errMessage = new ErrorMessage();
            errMessage.TopMost = true;
            providerName = providerName.ToLower();

            IOrderProvider provider;
            try {
                provider = GetProviderByName(providerName);
                if (provider is null) return;
            } catch (InvalidOperationException e) {
                errMessage.SetError($"Failed to get order provider '{providerName}'", e.Message, e.ToString());
                errMessage.ShowDialog();
                return;
            }

            Order order;
            try {
                order = provider.LoadCurrentOrder();
                if (order == null) throw new InvalidOperationException("No data was read");
            } catch (Exception e) {
                errMessage.SetError($"Failed to read order", e.Message, e.ToString());
                errMessage.ShowDialog();
                return;
            }

            Excel.Application app = ExcelDnaUtil.Application as Excel.Application;

            Worksheet outputsheet;

            try {
                outputsheet = app.Worksheets["Order"];
            } catch (Exception e) {
                errMessage.SetError($"Could not write order to worksheet", "Output sheet not found", "A properly formatted worksheet named 'Order' is required.\n----------------------------\n" + e.ToString());
                errMessage.ShowDialog();
                return;
            }

            try {
                OrderSink.WriteToSheet(outputsheet, order);
            } catch (Exception e) {
                errMessage.SetError("Failed to write order to sheet", e.Message, e.ToString());
                errMessage.ShowDialog();
                return;
            }

        }

        public static void PrintLabel(int line, int copies) {

            Worksheet dataSheet = ((Excel.Application)ExcelDnaUtil.Application).ActiveWorkbook.Sheets["Order"];

            var orderSource = dataSheet.Range["OrderSource"].Value2?.ToString().ToLower() ?? string.Empty;

            double height = dataSheet.Range["HeightCol"].Offset[line, 0].Value2;
            string heightStr = HelperFuncs.FractionalImperialDim(height);
            double width = dataSheet.Range["WidthCol"].Offset[line, 0].Value2;
            string widthStr = HelperFuncs.FractionalImperialDim(width);
            double depth = dataSheet.Range["DepthCol"].Offset[line, 0].Value2;
            string depthStr = HelperFuncs.FractionalImperialDim(depth);
            string size = $"{heightStr}H\" X {widthStr}W\" X {depthStr}D\"";

            try {
                if (orderSource == "hafele") {

                    HafeleLabelExport.PrintSingleHafeleLabel(
                            copies:         copies,
                            customerName:   dataSheet.Range["CustomerName"].Value2?.ToString() ?? "",
                            clientPO:       dataSheet.Range["OrderField_Value_5"].Value2?.ToString() ?? "",
                            hafelePO:       dataSheet.Range["OrderNumber"].Value2?.ToString() ?? "",
                            cfgNum:         dataSheet.Range["OrderField_Value_3"].Value2?.ToString() ?? "",
                            jobName:        dataSheet.Range["LevelCol"].Offset[line, 0].Value2?.ToString() ?? "",
                            qty:            dataSheet.Range["QtyCol"].Offset[line, 0].Value2?.ToString() ?? "",
                            lineNum:        line.ToString(),
                            size:           size,
                            message:        dataSheet.Range["NoteCol"].Offset[line,0].Value2?.ToString() ?? ""
                        );

                } else if (orderSource == "richlieu") {

                    RichelieuLabelExport.PrintSingleRichelieuLabel(
                            copies: copies,
                            jobName: dataSheet.Range["OrderField_Value_5"].Value2?.ToString() ?? "",
                            orderNum: dataSheet.Range["OrderNumber"].Value2?.ToString() ?? "",
                            size: size,
                            qty: dataSheet.Range["QtyCol"].Offset[line, 0].Value2?.ToString() ?? "",
                            description: dataSheet.Range["DescriptionCol"].Offset[line, 0].Value2?.ToString() ?? "",
                            richOrder: dataSheet.Range["OrderField_Value_1"].Value2?.ToString() ?? "",
                            note: dataSheet.Range["NoteCol"].Offset[line, 0].Value2?.ToString() ?? "",
                            lineNum: line.ToString()
                        );

                } else {

                    OTLabelExport.PrintSingleOTLabel(
                            copies: copies,
                            customerName: dataSheet.Range["CustomerName"].Value2?.ToString() ?? "",
                            size: size,
                            qty: dataSheet.Range["QtyCol"].Offset[line, 0].Value2?.ToString() ?? "",
                            orderNumber: dataSheet.Range["OrderNumber"].Value2?.ToString() ?? "",
                            lineNum: line.ToString(),
                            note: dataSheet.Range["NoteCol"].Offset[line, 0].Value2?.ToString() ?? "",
                            jobName: dataSheet.Range["OrderField_Value_1"].Value2?.ToString() ?? ""
                        );

                }
            } catch {
                System.Windows.Forms.MessageBox.Show("Error occurred printing single label");
            }


        }

        private static IOrderProvider GetProviderByName(string providerName) {
            string filepath = "";
            switch (providerName) {
                case "allmoxy":
                    filepath = ChooseFile();
                    if (filepath is null) return null;
                    return new AllmoxyOrderProvider(filepath);
                case "hafele":
                    filepath = ChooseFile();
                    if (filepath is null) return null;
                    return new HafeleDBOrderProvider(filepath);
                case "richelieu":
                    string input = Interaction.InputBox("Enter Richelieu web number of order to process", "Web Number", "", 0, 0);
                    if (input.Equals("")) return null;
                    return new RichelieuExcelDBOrderProvider(input);
                default:
                    throw new InvalidOperationException($"Unknown order provider '{providerName}'");
            }
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
