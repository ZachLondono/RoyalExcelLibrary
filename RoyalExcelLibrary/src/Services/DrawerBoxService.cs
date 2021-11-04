  using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.DAL;
using RoyalExcelLibrary.Providers;

using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Hosting;
using System.Threading.Tasks;
using System.Threading;
using RoyalExcelLibrary.DAL.Repositories;
using System.Data;
using RoyalExcelLibrary.Models.Products;

using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using RoyalExcelLibrary.ExportFormat;
using RoyalExcelLibrary.Models.Options;
using System.Collections;
using RoyalExcelLibrary.ExportFormat.CadCode;

namespace RoyalExcelLibrary.Services {
    public class DrawerBoxService : IProductService {


        public readonly IJobRepository JobRepository; // TODO replace this with a private field, use DI so that I don't need to get the instance from here
        private readonly IDrawerBoxRepository _drawerBoxRepository;
        private readonly IDbConnection _connection;

        private readonly ICutListFormat _stdCutlistFormat; 
        private readonly ICutListFormat _uboxCutlistFormat;
        private readonly CadCodeExport _cadExport;

        public DrawerBoxService(IDbConnection dbConnection) {
            _connection = dbConnection;
            JobRepository = new JobRepository(dbConnection);
            _drawerBoxRepository = new DrawerBoxRepository(dbConnection);
            
            _stdCutlistFormat = new StdCutListFormat();
            _uboxCutlistFormat = new UBoxCutListFormat();
            _cadExport = new CadCodeExport();
        }

        // <summary>
        // Stores the job in the current excel workbook in the job database and tracks the material it requires
        // </summar>
		public Order StoreCurrentOrder(Order order) {

            Job job = JobRepository.Insert(order.Job);
            order.Job.Id = job.Id;

            int count = 0;
            foreach (Product product in order.Products) {
                if (product is DrawerBox) {
                    DrawerBox drawerBox = (DrawerBox)product;
                    drawerBox.JobId = order.Job.Id;
                    _drawerBoxRepository.Insert(drawerBox);
                    count++;
                } 
            }

            return order;

        }

        public void SetOrderStatus(Order order, Status status) {
            order.Job.Status = status;
            JobRepository.Update(order.Job);
        }

        public Excel.Worksheet[] GenerateCutList(Order order, Excel.Workbook workbook) {

            Excel.Worksheet WriteCutlist(string worksheetname, IEnumerable<string[,]> seperatedBoxes, ICutListFormat cutListFormat) {

                if (seperatedBoxes.Count() == 0) return null;

                Excel.Worksheet outputsheet;

                try {
                    outputsheet = workbook.Worksheets[worksheetname];
                    outputsheet.Cells.Clear();
				} catch (COMException) {
                    outputsheet = workbook.Worksheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
                    outputsheet.Name = worksheetname;
				}

                Excel.Range header_rng = cutListFormat.WriteOrderHeader(order, outputsheet);
                Excel.Range cutlist_rng = cutListFormat.WriteOrderParts(seperatedBoxes, outputsheet, header_rng.Rows.Count + 1, 1);

                int startCol = header_rng.Column > cutlist_rng.Column ? cutlist_rng.Column : header_rng.Column;
                int startRow = header_rng.Row > cutlist_rng.Row ? cutlist_rng.Row : header_rng.Row;
                int endRow = startRow + header_rng.Rows.Count + cutlist_rng.Rows.Count;
                int endCol = startCol + (header_rng.Columns.Count > cutlist_rng.Columns.Count ? header_rng.Columns.Count : cutlist_rng.Columns.Count) - 1;

                Excel.Range print_rng = outputsheet.Range[outputsheet.Cells[startRow, startCol], outputsheet.Cells[endRow, endCol]];
                outputsheet.PageSetup.PrintArea = print_rng.Address;
                outputsheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

                return outputsheet;

            }

            List<DrawerBox> boxes = new List<DrawerBox>();
            foreach (var product in order.Products)
                boxes.Add((DrawerBox)product);

            // Sort by accending heights, then descending widths, then descending depths
            var sorted_boxes = boxes.OrderBy(b => b.Depth)
                                    .OrderByDescending(b => b.Width)
                                    .OrderByDescending(b => b.Height)
                                    .OrderBy(b => b is UDrawerBox);

            Excel.Worksheet std = WriteCutlist("CutList", AllParts(sorted_boxes), _stdCutlistFormat);
            Excel.Worksheet manual = WriteCutlist("Manual CutList", SimilarParts(sorted_boxes, DBPartType.Side), _stdCutlistFormat);
            if (!(manual is null)) manual.Range["H:H"].EntireColumn.Hidden = true; // Hides the Line# column
            Excel.Worksheet bottom = WriteCutlist("Bottom CutList", SimilarParts(sorted_boxes, DBPartType.Bottom), _stdCutlistFormat);
            if (!(bottom is null)) bottom.Range["H:H"].EntireColumn.Hidden = true;  // Hides the Line# column
            Excel.Worksheet ubox = null;
            if (sorted_boxes.Any(box => box is UDrawerBox)) {
                ubox = WriteCutlist("UBox CutList", UBoxParts(sorted_boxes), _uboxCutlistFormat);
                _cadExport.ExportOrder(order, $"R:\\DB ORDERS\\UBox Bottoms\\{order.Number}-UBoxs.csv");
            }

            return new Excel.Worksheet[] { std, bottom, manual, ubox};

        }

        public Excel.Worksheet GenerateConfirmation(Order order, Excel.Workbook outputBook) {
            throw new NotImplementedException();
        }

        public Excel.Worksheet GenerateInvoice(Order order, Excel.Workbook outputBook) {
            InvoiceExport invoiceExp = new InvoiceExport();

            ExportData data = new ExportData {
                SupplierName = order.Job.JobSource.ToLower().Equals("richelieu") ? "Royal Cabinet Co." : "Metro Drawer Boxes",
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

            return invoiceExp.ExportOrder(order, data, outputBook);
        }

        public Excel.Worksheet GeneratePackingList(Order order, Excel.Workbook outputBook) {
            PackingListExport packingListExp = new PackingListExport();

            ExportData data = new ExportData {
                SupplierName = order.Job.JobSource.ToLower().Equals("richelieu") ? "Royal Cabinet Co." : "Metro Drawer Boxes",
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

            return packingListExp.ExportOrder(order, data, outputBook);
        }

        private IEnumerable<string[,]> AllParts(IEnumerable<DrawerBox> boxes) {

            List<string[,]> formated = new List<string[,]>();

            int lineNum = 1;    // The line number of the part - in relation to the entire cutlist
            int partNum = 1;    // The line number of the part within the specific box
            foreach (DrawerBox box in boxes) {

                string height = HelperFuncs.FractionalImperialDim(box.Height);
                string width = HelperFuncs.FractionalImperialDim(box.Width);
                string depth = HelperFuncs.FractionalImperialDim(box.Depth);
                string sizeStr = $"{height}\"Hx{width}\"Wx{depth}\"D";

                IEnumerable<Part> parts = box.GetParts();

				string[,] part_rows = new string[parts.Count(), 9];

                string comm_1 = "";
                if (box.ScoopFront) comm_1 += "Scoop Front";
                if (box.Logo && box.ScoopFront) comm_1 += " | "; 
                if (box.Logo) comm_1 += "Logo";

                string comm_2 = "";
                if (box.PostFinish) comm_2 += "Post Finish";
                if (box.PostFinish && box.ClipsOption != Clips.No_Clips) comm_2 += " | ";
                if (box.ClipsOption != Clips.No_Clips) comm_2 += $"Clips: {box.ClipsOption}";

                string comm_3 = "";
                if (box is UDrawerBox) comm_3 += "UBox";
                if (box.NotchOption != UndermountNotch.No_Notch)
                    comm_3 += (comm_3.Length > 0 ? "\n" : "") + $"{box.NotchOption}";
                if (box.InsertOption != Insert.No_Insert)
                    comm_3 += (comm_3.Length > 0 ? "\n" : "") + $"Insert: {box.InsertOption}";
                if (box.MountingHoles)
                    comm_3 += (comm_3.Length > 0 ? "\n" : "") + $"Mounting Holes";

                partNum = 1;
                foreach (Part part in parts) {
                    part_rows[partNum - 1,0] = $"{box.LineNumber}";
                    part_rows[partNum - 1,1] = part.CutListName;
                    part_rows[partNum - 1, 2] = partNum == 1 ? comm_1 : partNum == 2 ? comm_2 :  partNum == 3 ? comm_3 : ""; 
                    part_rows[partNum - 1,3] = $"{part.Qty}";
                    part_rows[partNum - 1,4] = $"{Math.Round(part.Width,0)}";
                    part_rows[partNum - 1,5] = $"{Math.Round(part.Length, 0)}";
                    part_rows[partNum - 1,6] = MaterialCode(part.Material);
                    part_rows[partNum - 1,7] = $"{lineNum++}";
                    part_rows[partNum - 1,8] = sizeStr;

                    partNum++;
                }
                formated.Add(part_rows);
            }

            return formated;

		}

        private IEnumerable<string[,]> SimilarParts(IEnumerable<DrawerBox> boxes, DBPartType partType) {
            
            // Map a front to to the number of scoop fronts it has
            Dictionary<Part, int> scoopFronts = new Dictionary<Part, int>();
            List<(DrawerBoxPart, int)> parts = new List<(DrawerBoxPart, int)>();

            foreach (var box in boxes) {
                if (partType is DBPartType.Bottom && box is UDrawerBox) continue;
                foreach (var part in box.GetParts()) {

                    parts.Add(((DrawerBoxPart)part, box.LineNumber));

                    if (box.ScoopFront && part.CutListName.Contains("Front"))
                        scoopFronts.Add(part, box.Qty);    

                }
            }

            var filtered_parts = parts.Where(p => p.Item1.PartType == partType)
                                        .OrderByDescending(p => p.Item1.Length)
                                        .OrderByDescending(p => p.Item1.Width);

            // Map a part to a string with all the cab numbers in it
            Dictionary<Part, (string, int)> unique_parts = new Dictionary<Part, (string, int)>();

            foreach ((Part, int) item in filtered_parts) {

                Part part = item.Item1;
                bool match_found = false;
                foreach (var unique in unique_parts) {
                    var unique_part = unique.Key;
                    if (unique_part.Material == part.Material && unique_part.Width == part.Width && unique_part.Length == part.Length) {
                        unique_part.Qty += part.Qty;
                        match_found = true;
                        int scoopCount = unique_parts[unique_part].Item2;
                        if (scoopFronts.ContainsKey(item.Item1)) scoopCount += scoopFronts[part];
                        unique_parts[unique_part] = (unique.Value.Item1 + ", " + item.Item2, scoopCount);
                        break;
					}
				}

                if (!match_found) {
                    int scoopCount = 0;
                    if (scoopFronts.ContainsKey(part)) scoopCount = scoopFronts[part];
                    unique_parts.Add(part, ($"{item.Item2}", scoopCount));
                }

			}

            List<string[,]> part_rows = new List<string[,]>();

            int partnum = 0;
            foreach (var unique in unique_parts) {
                Part part = unique.Key;
                string boxnums = unique.Value.Item1;
                int scoopCount = unique.Value.Item2;

                string width = HelperFuncs.FractionalImperialDim(part.Width);
                string length = HelperFuncs.FractionalImperialDim(part.Length);

                string[,] part_row = new string[1,9];
                part_row[0, 0] = boxnums;
                part_row[0, 1] = part.CutListName;
                part_row[0, 2] = scoopCount == 0 ? "" : $"{scoopCount}x Scoop Fronts";
                part_row[0, 3] = $"{part.Qty}";
                part_row[0, 4] = $"{Math.Round(part.Width, 0)}";
                part_row[0, 5] = $"{Math.Round(part.Length, 0)}";
                part_row[0, 6] = MaterialCode(part.Material);
                part_row[0, 7] = $"{++partnum}";
               
                if (part is DrawerBoxPart && (part as DrawerBoxPart).PartType == DBPartType.Side)
                        part_row[0, 8] = $"{width}\"H x {length}\"L";
                else part_row[0, 8] = $"{width}\"W x {length}\"L";

                part_rows.Add(part_row);
            }

            return part_rows;

        }

        private IEnumerable<string[,]> UBoxParts(IEnumerable<DrawerBox> boxes) {

            var uboxes = boxes.Where(b => b is UDrawerBox);
            List<string[,]> formated = new List<string[,]>();

            int boxNum = 1;
            int lineNum = 1;
            foreach (UDrawerBox box in uboxes) {

                string height = HelperFuncs.FractionalImperialDim(box.Height);
                string width = HelperFuncs.FractionalImperialDim(box.Width);
                string depth = HelperFuncs.FractionalImperialDim(box.Depth);
                string sizeStr = $"{height}\"Hx{width}\"Wx{depth}\"D";

                IEnumerable<Part> parts = box.GetParts();

                string[,] part_rows = new string[parts.Count(), 9];

                int partnum = 1;
                foreach (Part part in parts) {

                    part_rows[partnum - 1, 0] = $"{boxNum}";
                    part_rows[partnum - 1, 1] = part.CutListName;
                    part_rows[partnum - 1, 2] = ""; // Comment
                    part_rows[partnum - 1, 3] = $"{part.Qty}";
                    part_rows[partnum - 1, 4] = $"{Math.Round(part.Width, 0)}";
                    part_rows[partnum - 1, 5] = $"{Math.Round(part.Length, 0)}";
                    part_rows[partnum - 1, 6] = MaterialCode(part.Material);
                    part_rows[partnum - 1, 7] = $"{lineNum++}";
                    part_rows[partnum - 1, 8] = sizeStr;

                    partnum++;
                }

                boxNum++;

                formated.Add(part_rows);
            }

            return formated;

        }

        // <summary>
        // String representation of material
        // </summary>
        private string MaterialCode(MaterialType material) {
            switch (material) {
                case MaterialType.EconomyBirch:
                    return "Birch FJ";
                case MaterialType.SolidBirch:
                    return "Birch CL";
                case MaterialType.SolidWalnut:
                    return "Walnut";
                case MaterialType.WhiteOak:
                    return "White Oak";
                case MaterialType.Plywood1_2:
                    return "Plywood 1/2";
                case MaterialType.Plywood1_4:
                    return "Plywood 1/4";
                default:
                    return "Unknown";
			}
		}

	}

}
