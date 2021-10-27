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

namespace RoyalExcelLibrary.Services {
    public class DrawerBoxService : IProductService {

        private readonly IJobRepository _jobRepository;
        private readonly IDrawerBoxRepository _drawerBoxRepository;
        private readonly IDbConnection _connection;

        private readonly ICutListFormat _stdCutlistFormat; 
        private readonly ICutListFormat _uboxCutlistFormat;

        public DrawerBoxService(IDbConnection dbConnection) {
            _connection = dbConnection;
            _jobRepository = new JobRepository(dbConnection);
            _drawerBoxRepository = new DrawerBoxRepository(dbConnection);
            
            _stdCutlistFormat = new StdCutListFormat();
            _uboxCutlistFormat = new UBoxCutListFormat();

        }

        // <summary>
        // Stores the job in the current excel workbook in the job database and tracks the material it requires
        // </summar>
		public Order StoreCurrentOrder(Order order) {

            Job job = _jobRepository.Insert(order.Job);
            order.Job.Id = job.Id;


            int count = 0;
            foreach (IProduct product in order.Products) {
                if (product is DrawerBox) {
                    DrawerBox drawerBox = (DrawerBox)product;
                    drawerBox.JobId = order.Job.Id;
                    _drawerBoxRepository.Insert(drawerBox);
                    count++;
                } 
            }

            return order;

        }

        public void GenerateCutList(Order order, Excel.Workbook workbook) {
                        
            void WriteCutlist(string worksheetname, IEnumerable<string[,]> seperatedBoxes, ICutListFormat cutListFormat) {

                if (seperatedBoxes.Count() == 0) return;

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
                int endCol = startCol + header_rng.Columns.Count + cutlist_rng.Columns.Count;

                Excel.Range print_rng = outputsheet.Range[outputsheet.Cells[startRow, startCol], outputsheet.Cells[endRow, endCol]];
                print_rng.Columns.AutoFit();
                print_rng.Rows.AutoFit();
                outputsheet.PageSetup.PrintArea = print_rng.Address;
                outputsheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            }

            List<DrawerBox> boxes = new List<DrawerBox>();
            foreach (var product in order.Products)
                boxes.Add((DrawerBox)product);

            // Sort by accending heights, then descending widths, then descending depths
            var sorted_boxes = boxes.OrderBy(b => b.Height)
                                    .OrderByDescending(b => b.Width)
                                    .OrderByDescending(b => b.Depth)
                                    .OrderBy(b => b is UDrawerBox);

            WriteCutlist("CutList", AllParts(sorted_boxes), _stdCutlistFormat);
            WriteCutlist("Bottom CutList", SimilarParts(sorted_boxes, DBPartType.Bottom), _stdCutlistFormat);
            WriteCutlist("Manual CutList", SimilarParts(sorted_boxes, DBPartType.Side), _stdCutlistFormat);
            WriteCutlist("UBox CutList", UBoxParts(sorted_boxes), _uboxCutlistFormat);

            // Mark job as released and update in database
            Job job = order.Job;
            job.Status = Status.Released;
            _jobRepository.Update(job);

        }

        public void GenerateConfirmation() {
			throw new NotImplementedException();
		}

		public void GenerateInvoice() {
			throw new NotImplementedException();
		}

		public void ConfirmOrder() {
			throw new NotImplementedException();
		}

        public void PayOrder() {
            throw new NotImplementedException();
        }

        private IEnumerable<string[,]> AllParts(IEnumerable<DrawerBox> boxes) {

            List<string[,]> formated = new List<string[,]>();

            int boxNum = 1;
            int lineNum = 1;
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

                int partnum = 1;
                foreach (Part part in parts) {
                    part_rows[partnum - 1,0] = $"{boxNum}";
                    part_rows[partnum - 1,1] = part.CutListName;
                    part_rows[partnum - 1, 2] = partnum == 1 ? comm_1 : partnum == 2 ? comm_2 :  partnum == 3 ? comm_3 : ""; 
                    part_rows[partnum - 1,3] = $"{part.Qty}";
                    part_rows[partnum - 1,4] = $"{Math.Round(part.Width,0)}";
                    part_rows[partnum - 1,5] = $"{Math.Round(part.Length, 0)}";
                    part_rows[partnum - 1,6] = MaterialCode(part.Material);
                    part_rows[partnum - 1,7] = $"{lineNum++}";
                    part_rows[partnum - 1,8] = sizeStr;

                    partnum++;
                }
                boxNum++;
                formated.Add(part_rows);
            }

            return formated;

		}

        private IEnumerable<string[,]> SimilarParts(IEnumerable<DrawerBox> boxes, DBPartType partType) {

            List<(DrawerBoxPart, int)> parts = new List<(DrawerBoxPart, int)>();
            int boxnum = 1;
            foreach (var box in boxes) {
                if (partType is DBPartType.Bottom && box is UDrawerBox) continue;
                foreach (var part in box.GetParts())
                    parts.Add(((DrawerBoxPart)part, boxnum));
                boxnum++;
            }

            var filtered_parts = parts.Where(p => p.Item1.PartType == partType)
                                        .OrderByDescending(p => p.Item1.Width)
                                        .OrderByDescending(p => p.Item1.Length);


            Dictionary<Part, string> unique_parts = new Dictionary<Part, string>();
            foreach ((Part, int) item in filtered_parts) {

                Part part = item.Item1;
                bool match_found = false;
                foreach (var unique in unique_parts) {
                    var unique_part = unique.Key;
                    if (unique_part.Material == part.Material && unique_part.Width == part.Width && unique_part.Length == part.Length) {
                        unique_part.Qty += part.Qty;
                        match_found = true;
                        unique_parts[unique_part] = unique.Value + ", " + item.Item2;
                        break;
					}
				}

                if (!match_found) unique_parts.Add(part, $"{item.Item2}");

			}

            List<string[,]> part_rows = new List<string[,]>();

            int partnum = 0;
            foreach (var unique in unique_parts) {
                Part part = unique.Key;
                string boxnums = unique.Value;

                string[,] part_row = new string[1,9];
                part_row[0, 0] = boxnums;
                part_row[0, 1] = part.CutListName;
                part_row[0, 2] = ""; // Comment
                part_row[0, 3] = $"{part.Qty}";
                part_row[0, 4] = $"{Math.Round(part.Width, 0)}";
                part_row[0, 5] = $"{Math.Round(part.Length, 0)}";
                part_row[0, 6] = MaterialCode(part.Material);
                part_row[0, 7] = $"{++partnum}";
                part_row[0, 8] = "";
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
