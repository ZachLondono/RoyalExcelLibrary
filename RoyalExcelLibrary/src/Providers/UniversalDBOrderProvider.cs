using RoyalExcelLibrary.Models;
using System;

using Microsoft.Office.Interop.Excel;
using RoyalExcelLibrary.Models.Products;
using RoyalExcelLibrary.Models.Options;

namespace RoyalExcelLibrary.Providers {
    public class UniversalDBOrderProvider : IExcelOrderProvider {

        public Application App { get; set; }
        private Worksheet _worksheet;

        public Order LoadCurrentOrder() {

            _worksheet = App.Worksheets["Order"];

            Job job = new Job {
                JobSource = _worksheet.Range["OrderSource"].Value2.ToString(),
                Name = "",
                CreationDate = DateTime.Today,
                GrossRevenue = 0
            };

            Order order;
            switch (job.JobSource.ToLower()) {
                case "hafele":
                    order = new HafeleOrder(job);
                    var hafOrder = order as HafeleOrder;
                    hafOrder.ProjectNumber =        _worksheet.Range["OrderField_Value_1"].Value2;
                    hafOrder.ProNumber =            _worksheet.Range["OrderField_Value_2"].Value2;
                    hafOrder.ConfigNumber =         _worksheet.Range["OrderField_Value_3"].Value2;
                    hafOrder.ClientAccountNumber =  _worksheet.Range["OrderField_Value_4"].Value2;
                    hafOrder.ClientPurchaseOrder =  _worksheet.Range["OrderField_Value_5"].Value2;
                    order.Job.Name = hafOrder.ClientPurchaseOrder; // TODO: get rid of job name
                    break;
                case "richelieu":
                    order = new RichelieuOrder(job);
                    var richOrder = order as RichelieuOrder;
                    richOrder.RichelieuNumber =     _worksheet.Range["OrderField_Value_1"].Value2;
                    richOrder.WebNumber =           _worksheet.Range["OrderField_Value_2"].Value2;
                    richOrder.ClientFirstName =     _worksheet.Range["OrderField_Value_3"].Value2;
                    richOrder.ClientLastName =      _worksheet.Range["OrderField_Value_4"].Value2;
                    richOrder.ClientPurchaseOrder = _worksheet.Range["OrderField_Value_5"].Value2;
                    richOrder.CustomerNum =         _worksheet.Range["OrderField_Value_6"].Value2;
                    order.Job.Name = richOrder.ClientPurchaseOrder; // TODO: get rid of job name
                    break;
                default:
                    order = new Order(job);
                    job.Name = _worksheet.Range["OrderField_Value_1"].Value2;
                    break;
            }

            order.Number = _worksheet.Range["OrderNumber"].Value2.ToString();
            order.SubTotal =    Convert.ToDecimal(_worksheet.Range["SubTotal"].Value2.ToString());
            order.Tax =         Convert.ToDecimal(_worksheet.Range["Tax"].Value2.ToString());
            order.ShippingCost =Convert.ToDecimal(_worksheet.Range["ShippingCost"].Value2.ToString());

            order.Customer = new Company {
                Name = _worksheet.Range["CustomerName"].Value2?.ToString() ?? "",
                Address = new ExportFormat.Address {
                    Line1 = _worksheet.Range["CustomerAddress1"].Value2?.ToString() ?? "",
                    Line2 = _worksheet.Range["CustomerAddress2"].Value2?.ToString() ?? "",
                    City = _worksheet.Range["CustomerCity"].Value2?.ToString() ?? "",
                    State = _worksheet.Range["CustomerState"].Value2?.ToString() ?? "",
                    Zip = _worksheet.Range["CustomerZip"].Value2?.ToString() ?? ""
                }
            };

            order.Vendor = new Company {
                Name = _worksheet.Range["VendorName"].Value2?.ToString() ?? "",
                Address = new ExportFormat.Address {
                    Line1 = _worksheet.Range["VendorAddress1"].Value2?.ToString() ?? "",
                    Line2 = _worksheet.Range["VendorAddress2"].Value2?.ToString() ?? "",
                    City = _worksheet.Range["VendorCity"].Value2?.ToString() ?? "",
                    State = _worksheet.Range["VendorState"].Value2?.ToString() ?? "",
                    Zip = _worksheet.Range["VendorZip"].Value2?.ToString() ?? ""
                }
            };

            order.Supplier = new Company {
                Name = _worksheet.Range["SupplierName"]?.Value2.ToString() ?? "",
                Address = new ExportFormat.Address {
                    Line1 = _worksheet.Range["SupplierAddress1"].Value2?.ToString() ?? "",
                    Line2 = _worksheet.Range["SupplierAddress2"].Value2?.ToString() ?? "",
                    City = _worksheet.Range["SupplierCity"].Value2?.ToString() ?? "",
                    State = _worksheet.Range["SupplierState"].Value2?.ToString() ?? "",
                    Zip = _worksheet.Range["SupplierZip"].Value2?.ToString() ?? ""
                }
            };


            var lineCol =       _worksheet.Range["LineCol"];
            var qtyCol =        _worksheet.Range["QtyCol"];
            var widthCol =      _worksheet.Range["WidthCol"];
            var heightCol =     _worksheet.Range["HeightCol"];
            var depthCol =      _worksheet.Range["DepthCol"];
            var dimACol =       _worksheet.Range["DimACol"];
            var dimBCol =       _worksheet.Range["DimBCol"];
            var dimCCol =       _worksheet.Range["DimCCol"];
            var materialCol =   _worksheet.Range["MaterialCol"];
            var bottomCol =     _worksheet.Range["BottomCol"];
            var notchCol =      _worksheet.Range["NotchCol"];
            var insertCol =     _worksheet.Range["InsertCol"];
            var clipCol =       _worksheet.Range["ClipCol"];
            var mountingHolesCol = _worksheet.Range["MountingHolesCol"];
            var finishCol =     _worksheet.Range["FinishCol"];
            var scoopCol =      _worksheet.Range["ScoopCol"];
            var logoCol =      _worksheet.Range["LogoCol"];
            var levelCol =      _worksheet.Range["LevelCol"];
            var noteCol =       _worksheet.Range["NoteCol"];
            var nameCol =       _worksheet.Range["NameCol"];
            var descriptionCol =_worksheet.Range["DescriptionCol"];
            var unitPriceCol =  _worksheet.Range["UnitPriceCol"];

            int offset = 1;
            while (lineCol.Offset[offset, 0] != null && !string.IsNullOrEmpty(lineCol.Offset[offset, 0].Value2?.ToString() ?? "")) {

                DrawerBox box;
                if (dimACol.Offset[offset, 0] != null && !string.IsNullOrEmpty(dimACol.Offset[offset, 0].Value2?.ToString() ?? "")) {
                    box = new UDrawerBox();
                    (box as UDrawerBox).A = dimACol.Offset[offset, 0]?.Value2 ?? 0;
                    (box as UDrawerBox).B = dimBCol.Offset[offset, 0]?.Value2 ?? 0;
                    (box as UDrawerBox).C = dimCCol.Offset[offset, 0]?.Value2 ?? 0;
                } else box = new DrawerBox();

                box.LineNumber =    Convert.ToInt32(lineCol.Offset[offset, 0].Value2);
                box.Qty =           Convert.ToInt32(qtyCol.Offset[offset,0].Value2);
                box.Width =         Convert.ToDouble(widthCol.Offset[offset, 0].Value2);
                box.Height =        Convert.ToDouble(heightCol.Offset[offset, 0].Value2);
                box.Depth =         Convert.ToDouble(depthCol.Offset[offset, 0].Value2);
                box.SideMaterial =  ParseMaterial(materialCol.Offset[offset, 0].Value2);
                box.BottomMaterial = ParseMaterial(bottomCol.Offset[offset, 0].Value2);
                box.NotchOption =   ParseNotch(notchCol.Offset[offset, 0].Value2);
                box.InsertOption=   insertCol.Offset[offset, 0].Text;
                box.ClipsOption =   ParseClips(clipCol.Offset[offset, 0].Value2);
                box.MountingHoles = mountingHolesCol.Offset[offset, 0].Value2.Equals("Yes") ? true : false;
                box.PostFinish =    finishCol.Offset[offset, 0].Value2.Equals("Yes") ? true : false;
                box.ScoopFront =    scoopCol.Offset[offset, 0].Value2.Equals("Yes") ? true : false;
                box.Logo =          logoCol.Offset[offset, 0].Value2.Equals("Yes") ? true : false;
                box.LevelName =     levelCol.Offset[offset, 0].Value2?.ToString() ?? "";
                box.ProductName =   nameCol.Offset[offset,0].Value2?.ToString() ?? "";
                box.ProductDescription = descriptionCol.Offset[offset,0].Value2?.ToString() ?? "";
                box.Note =          noteCol.Offset[offset, 0].Value2?.ToString() ?? "";
                box.UnitPrice =     Convert.ToDecimal(unitPriceCol.Offset[offset, 0].Value2);

                order.AddProduct(box);

                offset++;

            }

            return order;

        }

        private MaterialType ParseMaterial(string name) {
            switch (name) {
                case "BlackMela1_2":
                    return MaterialType.BlackMela1_2;
                case "BlackMela1_4":
                    return MaterialType.BlackMela1_4;
                case "WhiteMela1_2":
                    return MaterialType.WhiteMela1_2;
                case "WhiteMela1_4":
                    return MaterialType.WhiteMela1_4;
                case "Plywood1_2":
                    return MaterialType.Plywood1_2;
                case "Plywood1_4":
                    return MaterialType.Plywood1_4;
                case "EconomyBirch":
                    return MaterialType.EconomyBirch;
                case "SolidBirch":
                    return MaterialType.SolidBirch;
                case "SolidWalnut":
                    return MaterialType.SolidWalnut;
                case "WhiteOak":
                    return MaterialType.WhiteOak;
                case "HybridBirch":
                    return MaterialType.HybridBirch;
                default:
                    return MaterialType.Unknown;
            }   
        }

        private UndermountNotch ParseNotch(string name) {
            switch (name) {
                case "No_Notch":
                    return UndermountNotch.No_Notch;
                case "Std_Notch":
                    return UndermountNotch.Std_Notch;
                case "Notch_828":
                    return UndermountNotch.Notch_828;
                case "Wide_Notch":
                    return UndermountNotch.Wide_Notch;
                case "Front_Back":
                    return UndermountNotch.Front_Back;
                default:
                    return UndermountNotch.Unknown;
            }
        }

        private Clips ParseClips(string name) {
            switch (name) {
                case "Hafele":
                    return Clips.Hafele;
                case "Blum":
                    return Clips.Blum;
                case "Hettich":
                    return Clips.Hettich;
                case "No_Clips":
                    return Clips.No_Clips;
                case "Richelieu":
                    return Clips.Richelieu;
                default:
                    return Clips.Unknown;
            }
        }

    }
}
