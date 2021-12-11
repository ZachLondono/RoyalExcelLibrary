using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using RoyalExcelLibrary.Providers;
using RoyalExcelLibrary.Services;
using RoyalExcelLibrary.Models;
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
using Label = RoyalExcelLibrary.Services.Label;
using System.Data.OleDb;
using System.Data;
using RoyalExcelLibrary.Models.Products;

namespace RoyalExcelLibrary {
	public class ExcelLibrary {

#if DEBUG
            public const string ConnectionString = "Provider=Microsoft.ace.OLEDB.12.0; Data Source='R:\\DB ORDERS\\RoyalExcelLibrary\\TestData.accdb';";
#else
            public const string ConnectionString = "Provider=Microsoft.ace.OLEDB.12.0; Data Source='R:\\DB ORDERS\\RoyalExcelLibrary\\Data.accdb';";
#endif

        public static void DrawerBoxProcessor(string format, bool trackjob, bool generateCutLists, bool printLabels, bool printCutlists, bool generatePackingList, bool printPackingList, bool generateInvoice, bool printInvoice, bool emailInvoice) {

#if DEBUG
            MessageBox.Show($"Starting in Debug Mode\n Using database: '{ConnectionString}'");
#endif

            ErrorMessage errMessage = new ErrorMessage {
                TopMost = true
            };

            Excel.Application app = ExcelDnaUtil.Application as Excel.Application;
            Worksheet initialWorksheet = app.ActiveSheet;

            bool printbol = false;
            try {
                Shape bolCheckbox = app.Worksheets["Order"].Shapes.Item("PrintBOL");
                if (!(bolCheckbox is null))
                    printbol = (bolCheckbox.OLEFormat.Object.Value == 1);
            } catch {
                Console.WriteLine("No bol checkbox found");
            }

            IOrderProvider provider = GetProviderByName(format);
            if (provider is RichelieuExcelDBOrderProvider) {
                string webnumber = Interaction.InputBox("Enter Richelieu web number of order to process", "Web Number", "", 0, 0);
                if (webnumber.Equals("")) return;
                (provider as RichelieuExcelDBOrderProvider).DownloadOrder(webnumber);
            } else if (provider is IExcelOrderProvider) {
                (provider as IExcelOrderProvider).App = app;
            } else if (provider is IFileOrderProvider) {
                string path = ChooseFile();
                if (string.IsNullOrEmpty(path)) return;
                (provider as IFileOrderProvider).FilePath = path;
            }

            Order order;
            try {
                app.ScreenUpdating = false;
                order = provider.LoadCurrentOrder();
                app.ScreenUpdating = true;

                if (order is RichelieuOrder) {
                    if ((order as RichelieuOrder).Rush) {
                        MessageBox.Show("This order is a 3-Day Rush", "Rush Order", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

            } catch (Exception e) {
                app.ScreenUpdating = true;
                errMessage.SetError("Error Loading Order", e.Message, e.ToString());
                errMessage.ShowDialog();
                return;
            }

            Dictionary<string, Worksheet> printQueue = new Dictionary<string, Worksheet>();

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

            if (trackjob) {
                switch (order.Job.JobSource.ToLower()) {
                    case "hafele":
                        new HafeleGoogleSheetExport().ExportOrder(order);
                        break;
                    case "richelieu":
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

                using (OleDbConnection dbConnection = new OleDbConnection(ConnectionString)) {

                    int jobId = -1;

                    dbConnection.Open();
                    try { 
                        // Track Job name
                        jobId = TrackJobInDB(dbConnection, order.Number, DateTime.Today, order.SubTotal + order.ShippingCost, order.Vendor.Name);
                    } catch (Exception e) {
                        Console.WriteLine("Error tracking job");
                        Console.WriteLine(e);
                    }

                    // Track drawers used in order
                    try {
                        TrackItemsInDB(dbConnection, jobId, order.Products);
                    } catch (Exception e) {
                        Console.WriteLine("Error tracking items");
                        Console.WriteLine(e);
                    }

                    // Track optimized material used in order
                    try {
                        TrackMaterialInDB(dbConnection, jobId, order.Products);
                    } catch (Exception e) {
                        Console.WriteLine("Error tracking material use");
                        Console.WriteLine(e);
                    }

                    try {

                        string vendorName = "";
                        string customerName = order.Customer.Name;
                        Address billingAddress = order.Customer.Address;
                        switch (order.Job.JobSource.ToLower()) {
                            case "richelieu":
                                vendorName = order.Vendor.Name;
                                customerName = "Richelieu";
                                billingAddress = order.Vendor.Address;
                                break;
                            case "hafele":
                                vendorName = order.Vendor.Name;
                                customerName = "Hafele";
                                billingAddress = order.Vendor.Address;
                                break;
                            case "ot":
                            case "on track":
                                vendorName = "OT";
                                break;
                            case "royal":
                                vendorName = "Royal Cabinet Co.";
                                break;
                            case "allmoxy":
                                vendorName = "Metro Cabinet Parts";
                                break;
                            default:
                                break;
                        }


                        TrackInvoiceInDB(dbConnection,
                            customer:           customerName,
                            transactionDate:    DateTime.Today,
                            PONumber:           order.Job.Name,
                            refNumber:          order.Number,
                            item:               "Drawer Boxes",
                            description:        "Drawer Boxes",
                            price:              order.SubTotal,
                            vendor:             vendorName,
                            billingAddress:     billingAddress);
                    } catch (Exception e) {
                        Console.WriteLine("Error tracking invoice information");
                        Console.WriteLine(e);
                    }
                }

            }

            IProductService productService = new DrawerBoxService();

            if (generateCutLists) {

                try {
                    app.ScreenUpdating = false;
                    app.DisplayAlerts = false;
                    var cutlists = productService.GenerateCutList(order, app.ActiveWorkbook, errMessage);

                    if (printCutlists)
                        foreach (var cutlist in cutlists) {
                            printQueue.Add(cutlist.Key, cutlist.Value);
                        }

                    if (order.Job.JobSource.ToLower().Equals("hafele")) {
                        BOLExport bolExpt = new BOLExport();
                        var bol = bolExpt.ExportOrder(order, app.ActiveWorkbook);

                        if (printbol) {
                            printQueue.Add("bol",bol);
                            PrintBOLLabel();
                        }
                    }
                    app.DisplayAlerts = true;
                    app.ScreenUpdating = true;
                } catch (Exception e) {
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    errMessage.SetError("Error While Cut Listing", e.Message, e.ToString());
                    errMessage.ShowDialog();
                }

            }

            if (generatePackingList) {

                try {

                    app.ScreenUpdating = false;
                    app.DisplayAlerts = false;
                    Worksheet packingList = productService.GeneratePackingList(order, app.ActiveWorkbook, errMessage);
                    app.DisplayAlerts = true;
                    app.ScreenUpdating = true;

                    if (printPackingList)
                        printQueue.Add("packing",packingList);

                } catch (Exception e) {
                    app.ScreenUpdating = true;
                    app.DisplayAlerts = true;
                    errMessage.SetError("Error While Generating/Printing Packing List", e.Message, e.ToString());
                    errMessage.ShowDialog();
                }

            }

            if (generateInvoice) {

                try {

                    InvoiceExport invoiceExp = new InvoiceExport();

                    app.ScreenUpdating = false;
                    app.DisplayAlerts = false;
                    Worksheet invoice = productService.GenerateInvoice(order, app.ActiveWorkbook, errMessage);
                    app.DisplayAlerts = true;
                    app.ScreenUpdating = true;

                    if (printInvoice) {
                        printQueue.Add("invoice", invoice);
                    } 

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
                    app.DisplayAlerts = true;
                    errMessage.SetError("Error While Generating/Printing Invoice", e.Message, e.ToString());
                    errMessage.ShowDialog();
                }

            }

            if (printLabels) {

                try {

                    ILabelExport labelExport;
                    switch (order.Job.JobSource.ToLower()) {
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

            string[] printOrder = {
                "manual",
                "standard",
                "ubox",
                "packing",
                "invoice",
                "bol",
                "bottom"
            };

            try {
                if (!printerInstalled)
                    throw new InvalidOperationException($"The printer {printerName} could not be accessed");

                foreach (string sheetName in printOrder) {
                    if (printQueue.ContainsKey(sheetName)) {
                        var sheet = printQueue[sheetName];
                        if (sheet is null) continue;

                        int copies = 1;
                        if (sheetName == "packing" && order.Job.JobSource.ToLower() == "richelieu")
                            copies = 3;

                        printQueue[sheetName].PrintOutEx(ActivePrinter: printerName, Copies: copies);
                    }
                }
            } catch (Exception e) {
                errMessage.SetError("Error While Printing", e.Message, e.ToString());
                errMessage.ShowDialog();
            }

            initialWorksheet.Select();

            errMessage.Dispose();

        }

        public static void PrintBOLLabel() {
            DymoLabelService aduieLabelService = new DymoLabelService("R:\\DB ORDERS\\Labels\\Duie Pyle notice.label");
            Label aduielabel = aduieLabelService.CreateLabel();
            aduieLabelService.AddLabel(aduielabel, 1);
            aduieLabelService.PrintLabels();
        }

        /// <summary>
        /// Loads an order from provider and writes it to the excel worksheet
        /// </summary>
        /// <param name="providerName"></param>
        public static void LoadOrder(string providerName) {

            ErrorMessage errMessage = new ErrorMessage {
                TopMost = true
            };
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

            Excel.Application app = ExcelDnaUtil.Application as Excel.Application;

            if (provider is RichelieuExcelDBOrderProvider) {
                string webnumber = Interaction.InputBox("Enter Richelieu web number of order to process", "Web Number", "", 0, 0);
                if (webnumber.Equals("")) return;
                (provider as RichelieuExcelDBOrderProvider).DownloadOrder(webnumber);
            } else if (provider is IExcelOrderProvider) {
                (provider as IExcelOrderProvider).App = app;
            } else if (provider is IFileOrderProvider) {
                string path = ChooseFile();
                if (string.IsNullOrEmpty(path)) return;
                (provider as IFileOrderProvider).FilePath = path;
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

        public static void PostOrderToGoogleSheets() {

            VendorSelector vendorSelector = new VendorSelector {
                TopMost = true
            };

            var result = vendorSelector.ShowDialog();

            if (result != DialogResult.OK) {
                return;
            }

            string vendor = vendorSelector.GetSelected();

            var provider = new UniversalDBOrderProvider();
            provider.App = ExcelDnaUtil.Application as Excel.Application;

            var order = provider.LoadCurrentOrder();

            switch (vendor.ToLower()) {
                case "hafele":
                    new HafeleGoogleSheetExport().ExportOrder(order);
                    break;
                case "richelieu":
                    new RichelieuGoogleSheetExport().ExportOrder(order);
                    break;
                case "on track":
                    new OTGoogleSheetExport().ExportOrder(order);
                    break;
                case "metro":
                    new MetroGoogleSheetExport().ExportOrder(order);
                    break;
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

                } else if (orderSource == "richelieu") {

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
                MessageBox.Show("Error occurred printing single label");
            }


        }

        public static IOrderProvider GetProviderByName(string providerName) {
            switch (providerName.ToLower()) {
                case "ot":
                    return new OTDBOrderProvider();
                case "hafele":
                    return new HafeleDBOrderProvider();
                case "richelieu":
                    return new RichelieuExcelDBOrderProvider();
                case "allmoxy":
                    return new AllmoxyOrderProvider();
                case "loaded":
                    return new UniversalDBOrderProvider();
                default:
                    throw new ArgumentException("Unknown provider provider '{providerName}'");
            }
        }

		public static string ChooseFile() {
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

        private static int TrackJobInDB(OleDbConnection connection, string name, DateTime creationDate, decimal grossRevenue, string vendor) {
            using (OleDbCommand command = new OleDbCommand()) {

                command.Connection = connection;
                command.CommandType = CommandType.Text;

                command.CommandText = "INSERT INTO Jobs ([JobName], [CreationDate], [GrossRevenue], [Vendor]) VALUES (@name, @creationDate, @grossRevenue, @vendor);";

                command.Parameters.Add(new OleDbParameter("@name", OleDbType.VarChar)).Value = name;
                command.Parameters.Add(new OleDbParameter("@creationDate", OleDbType.Date)).Value = creationDate;
                command.Parameters.Add(new OleDbParameter("@grossRevenue", OleDbType.Decimal)).Value = grossRevenue;
                command.Parameters.Add(new OleDbParameter("@vendor", OleDbType.VarChar)).Value = vendor;

                command.ExecuteNonQuery();

                command.CommandText = "select @@IDENTITY from Jobs;";

                var reader = command.ExecuteReader();
                reader.Read();
                return reader.GetInt32(0);
            }
        }

        private static void TrackItemsInDB(OleDbConnection connection, int jobId, IEnumerable<Product> products) {


            foreach (Product product in products) {

                using (OleDbCommand command = new OleDbCommand()) {

                    command.Connection = connection;
                    command.CommandType = CommandType.Text;

                    var drawerbox = (DrawerBox)product;

                    command.CommandText = @"INSERT INTO DrawerBoxes ([Qty], [Height], [Width], [Depth], [SideMaterial], [BottomMaterial], [JobId])
                                            VALUES (@qty, @height, @width, @depth, @side, @bottom, @jobId);";

                    command.Parameters.Add(new OleDbParameter("@qty", OleDbType.Integer)).Value = drawerbox.Qty;
                    command.Parameters.Add(new OleDbParameter("@height", OleDbType.Double)).Value = drawerbox.Height;
                    command.Parameters.Add(new OleDbParameter("@width", OleDbType.Double)).Value = drawerbox.Width;
                    command.Parameters.Add(new OleDbParameter("@depth", OleDbType.Double)).Value = drawerbox.Depth;
                    command.Parameters.Add(new OleDbParameter("@side", OleDbType.VarChar)).Value = drawerbox.SideMaterial.ToString();
                    command.Parameters.Add(new OleDbParameter("@bottom", OleDbType.VarChar)).Value = drawerbox.BottomMaterial.ToString();
                    command.Parameters.Add(new OleDbParameter("@jobId", OleDbType.Integer)).Value = jobId;

                    command.ExecuteNonQuery();

                }

            }

        }

        private static void TrackMaterialInDB(OleDbConnection connection, int jobId, IEnumerable<Product> products) {

            foreach (Product product in products) {
                using (OleDbCommand command = new OleDbCommand()) {
                    command.Connection = connection;
                    command.CommandType = CommandType.Text;

                    var trackDate = DateTime.Today;

                    foreach (Part part in product.GetParts()) {                        

                        command.CommandText = @"INSERT INTO Parts ([Qty], [Width], [Length], [Thickness], [Material], [Timestamp], [JobId])
                                            VALUES (@qty, @width, @length, @thickness, @material, @timestamp, @jobId);";

                        command.Parameters.Add(new OleDbParameter("@qty", OleDbType.Integer)).Value = part.Qty;
                        command.Parameters.Add(new OleDbParameter("@width", OleDbType.Double)).Value = part.Width;
                        command.Parameters.Add(new OleDbParameter("@length", OleDbType.Double)).Value = part.Length;
                        command.Parameters.Add(new OleDbParameter("@thickness", OleDbType.Double)).Value = 0;
                        command.Parameters.Add(new OleDbParameter("@material", OleDbType.VarChar)).Value = part.Material.ToString();
                        command.Parameters.Add(new OleDbParameter("@timestamp", OleDbType.Date)).Value = trackDate;
                        command.Parameters.Add(new OleDbParameter("@jobId", OleDbType.Integer)).Value = jobId;

                        command.ExecuteNonQuery();

                    }

                }

            }

        }

        private static void TrackInvoiceInDB(OleDbConnection connection, string customer, DateTime transactionDate, string PONumber, string refNumber, string item, string description, decimal price, string vendor, Address billingAddress) {

            using (OleDbCommand command = new OleDbCommand()) {

                command.Connection = connection;
                command.CommandType = CommandType.Text;

                command.CommandText = @"INSERT INTO Invoices
                                        ([Customer], [TransactionDate], [PONumber], [RefNumber], [Item], [Description], [Price], [Status], [Vendor], [AddressLine1], [AddressLine2], [City], [State], [PostalCode], [Country])
                                        VALUES
                                        (@Customer, @TransactionDate, @PONumber, @RefNumber, @Item, @Description, @Price, @Status, @Vendor, @AddressLine1, @AddressLine2, @City, @State, @PostalCode, @Country)";

                command.Parameters.Add(new OleDbParameter("@Customer", OleDbType.VarChar)).Value = customer;
                command.Parameters.Add(new OleDbParameter("@TransactionDate", OleDbType.Date)).Value = transactionDate;
                command.Parameters.Add(new OleDbParameter("@PONumber", OleDbType.VarChar)).Value = PONumber;
                command.Parameters.Add(new OleDbParameter("@RefNumber", OleDbType.VarChar)).Value = refNumber;
                command.Parameters.Add(new OleDbParameter("@Item", OleDbType.VarChar)).Value = item;
                command.Parameters.Add(new OleDbParameter("@Description", OleDbType.VarChar)).Value = description;
                command.Parameters.Add(new OleDbParameter("@Price", OleDbType.Currency)).Value = price;
                command.Parameters.Add(new OleDbParameter("@Status", OleDbType.VarChar)).Value = "UnExported";
                command.Parameters.Add(new OleDbParameter("@Vendor", OleDbType.VarChar)).Value = vendor;
                command.Parameters.Add(new OleDbParameter("@AddressLine1", OleDbType.VarChar)).Value = billingAddress.Line1;
                command.Parameters.Add(new OleDbParameter("@AddressLine2", OleDbType.VarChar)).Value = billingAddress.Line2;
                command.Parameters.Add(new OleDbParameter("@City", OleDbType.VarChar)).Value = billingAddress.City;
                command.Parameters.Add(new OleDbParameter("@State", OleDbType.VarChar)).Value = billingAddress.State;
                command.Parameters.Add(new OleDbParameter("@PostalCode", OleDbType.VarChar)).Value = billingAddress.Zip;
                command.Parameters.Add(new OleDbParameter("@Country", OleDbType.VarChar)).Value = "USA";

                command.ExecuteNonQuery();

            }

        }

	}

}
