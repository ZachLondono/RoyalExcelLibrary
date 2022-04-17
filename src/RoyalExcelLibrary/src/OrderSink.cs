using System;
using System.Linq;
using System.Collections;

using Microsoft.Office.Interop.Excel;
using RoyalExcelLibrary.ExcelUI.Models;
using RoyalExcelLibrary.ExcelUI.Models.Products;

namespace RoyalExcelLibrary.ExcelUI {
    public class OrderSink {

        public static void WriteToSheet(Worksheet outputSheet, Order order) {

            try {
                outputSheet.Range["ClearArea_1"].Value2 = "";
                outputSheet.Range["ClearArea_2"].Value2 = "";
                outputSheet.Range["ClearArea_3"].Value2 = "";
                outputSheet.Range["ClearArea_4"].Value2 = "";
                outputSheet.Range["ClearArea_5"].Value2 = "";
                outputSheet.Range["ClearArea_6"].Value2 = "";
                outputSheet.Range["ClearArea_7"].Value2 = "";
                outputSheet.Range["ClearArea_8"].Value2 = "";
                outputSheet.Range["OrderComment"].Value2 = "";
                outputSheet.Range["ShippingInstructions"].Value2 = "";
            } catch (Exception e) {
                Console.WriteLine("Failed to clear ranges " + e.ToString());
            }

            var customer = order.Customer;
            outputSheet.Range["CustomerName"].Value2 =      customer.Name;
            outputSheet.Range["CustomerAddress1"].Value2 =  customer.Address?.Line1 ?? "";
            outputSheet.Range["CustomerAddress2"].Value2 =  customer.Address?.Line2 ?? "";
            outputSheet.Range["CustomerCity"].Value2 =      customer.Address?.City ?? "";
            outputSheet.Range["CustomerState"].Value2 =     customer.Address?.State ?? "";
            outputSheet.Range["CustomerZip"].Value2 =       customer.Address?.Zip ?? "";

            var vendor = order.Vendor;
            outputSheet.Range["VendorName"].Value2 =        vendor.Name;
            outputSheet.Range["VendorAddress1"].Value2 =    vendor.Address?.Line1 ?? "";
            outputSheet.Range["VendorAddress2"].Value2 =    vendor.Address?.Line2 ?? "";
            outputSheet.Range["VendorCity"].Value2 =        vendor.Address?.City ?? "";
            outputSheet.Range["VendorState"].Value2 =       vendor.Address?.State ?? "";
            outputSheet.Range["VendorZip"].Value2 =         vendor.Address?.Zip ?? "";

            var supplier = order.Supplier;
            outputSheet.Range["SupplierName"].Value2 =      supplier.Name;
            outputSheet.Range["SupplierAddress1"].Value2 =  supplier.Address?.Line1 ?? "";
            outputSheet.Range["SupplierAddress2"].Value2 =  supplier.Address?.Line2 ?? "";
            outputSheet.Range["SupplierCity"].Value2 =      supplier.Address?.City ?? "";
            outputSheet.Range["SupplierState"].Value2 =     supplier.Address?.State ?? "";
            outputSheet.Range["SupplierZip"].Value2 =       supplier.Address?.Zip ?? "";

            var orderNum =          outputSheet.Range["OrderNumber"];
            orderNum.Value2 =       order.Number.ToString();
            var orderSource =       outputSheet.Range["OrderSource"];
            orderSource.Value2 =    order.Job.JobSource.ToString();

            if (order.Rush) { 
                var rushRng =           outputSheet.Range["RushMessage"];
                rushRng.Value2 =        "Rush Order";
            }

            var orderField_1 = outputSheet.Range["OrderField_Key_1"];
            var orderFieldValue_1 = outputSheet.Range["OrderField_Value_1"];
            var orderField_2 = outputSheet.Range["OrderField_Key_2"];
            var orderFieldValue_2 = outputSheet.Range["OrderField_Value_2"];
            var orderField_3 = outputSheet.Range["OrderField_Key_3"];
            var orderFieldValue_3 = outputSheet.Range["OrderField_Value_3"];
            var orderField_4 = outputSheet.Range["OrderField_Key_4"];
            var orderFieldValue_4 = outputSheet.Range["OrderField_Value_4"];
            var orderField_5 = outputSheet.Range["OrderField_Key_5"];
            var orderFieldValue_5 = outputSheet.Range["OrderField_Value_5"];
            var orderField_6 = outputSheet.Range["OrderField_Key_6"];
            var orderFieldValue_6 = outputSheet.Range["OrderField_Value_6"];

            if (order is HafeleOrder) {

                var hafOrder = order as HafeleOrder;

                orderField_1.Value2 = "Project Number";
                orderFieldValue_1.Value2 = hafOrder.ProjectNumber;

                orderField_2.Value2 = "A Duie Pyle #";
                orderFieldValue_2.Value2 = hafOrder.ProNumber;

                orderField_3.Value2 = "Config Number";
                orderFieldValue_3.Value2 = hafOrder.ConfigNumber;

                orderField_4.Value2 = "Client Account";
                orderFieldValue_4.Value2 = hafOrder.ClientAccountNumber;

                orderField_5.Value2 = "Client PO";
                orderFieldValue_5.Value2 = hafOrder.ClientPurchaseOrder;

                outputSheet.Range["OrderSourceLink"].Value2 = hafOrder.SourceFile;

            } else if (order is RichelieuOrder) {

                var richOrder = order as RichelieuOrder;

                orderField_1.Value2 = "Richelieu #";
                orderFieldValue_1.Value2 = richOrder.RichelieuNumber;

                orderField_2.Value2 = "Web #";
                orderFieldValue_2.Value2 = richOrder.WebNumber;

                orderField_3.Value2 = "First Name";
                orderFieldValue_3.Value2 = richOrder.ClientFirstName;

                orderField_4.Value2 = "Last Name";
                orderFieldValue_4.Value2 = richOrder.ClientLastName;

                orderField_5.Value2 = "Client PO";
                orderFieldValue_5.Value2 = richOrder.ClientPurchaseOrder;

                orderField_6.Value2 = "Customer Number";
                orderFieldValue_6.Value2 = richOrder.CustomerNum;

            } else if (order is AllmoxyOrder) {

                var allmoxyOrder = order as AllmoxyOrder;

                orderField_1.Value2 = "Job Name";
                orderFieldValue_1.Value2 = order.Job.Name;

                outputSheet.Range["ShippingInstructions"].Value2 = allmoxyOrder.ShippingInstructions;

            }

            if (order.Job.JobSource.ToLower().Equals("allmoxy")) {
                outputSheet.Range["OrderSourceLink"].Value2 = $"https://metrodrawerboxes.allmoxy.com/orders/quote/{order.Number}/";
            }

            var range = outputSheet.Range["OrderComment"];
            if (!(range is null) && !string.IsNullOrEmpty(order.Comment)) {
                range.Value2 = order.Comment;
            }

            var subTotal =      outputSheet.Range["SubTotal"];
            subTotal.Value2 =   order.SubTotal.ToString();
            var tax =           outputSheet.Range["Tax"];
            tax.Value2 =        order.Tax.ToString();
            var shipping =      outputSheet.Range["ShippingCost"];
            shipping.Value2 =   order.ShippingCost.ToString();
            var total =         outputSheet.Range["TotalCost"];
            total.Value2 =      (order.SubTotal + order.Tax + order.ShippingCost).ToString();

            var lineCol =       outputSheet.Range["LineCol"];
            var qtyCol =        outputSheet.Range["QtyCol"];
            var widthCol =      outputSheet.Range["WidthCol"];
            var heightCol =     outputSheet.Range["HeightCol"];
            var depthCol =      outputSheet.Range["DepthCol"];
            var dimACol =       outputSheet.Range["DimACol"];
            var dimBCol =       outputSheet.Range["DimBCol"];
            var dimCCol =       outputSheet.Range["DimCCol"];
            var materialCol =   outputSheet.Range["MaterialCol"];
            var bottomCol =     outputSheet.Range["BottomCol"];
            var insertCol =     outputSheet.Range["InsertCol"];
            var notchCol =      outputSheet.Range["NotchCol"];
            var clipCol =       outputSheet.Range["ClipCol"];
            var mountingHolesCol = outputSheet.Range["MountingHolesCol"];
            var finishCol =     outputSheet.Range["FinishCol"];
            var scoopCol =      outputSheet.Range["ScoopCol"];
            var logoCol =       outputSheet.Range["LogoCol"];
            var levelCol =      outputSheet.Range["LevelCol"];
            var noteCol =       outputSheet.Range["NoteCol"];
            var nameCol =       outputSheet.Range["NameCol"];
            var descriptionCol =outputSheet.Range["DescriptionCol"];
            var unitPriceCol =  outputSheet.Range["UnitPriceCol"];
            var linkCol =       outputSheet.Range["LinkCol"];

            var boxes = order.Products.Where(p => p is DrawerBox).Cast<DrawerBox>().ToList();

            outputSheet.Range["BoxCount"].Value2 = boxes.Sum(b => b.Qty);

            int boxCount = boxes.Count;

            string[,] lines = new string[boxCount, 1];
            string[,] qtys = new string[boxCount, 1];
            string[,] widths = new string[boxCount, 1];
            string[,] heights = new string[boxCount, 1];
            string[,] depths = new string[boxCount, 1];
            string[,] materials = new string[boxCount, 1];
            string[,] bottoms = new string[boxCount, 1];
            string[,] notches = new string[boxCount, 1];
            string[,] inserts = new string[boxCount, 1];
            string[,] clips = new string[boxCount, 1];
            string[,] mountingHoles = new string[boxCount, 1];
            string[,] finishes = new string[boxCount, 1];
            string[,] scoops = new string[boxCount, 1];
            string[,] logos = new string[boxCount, 1];
            string[,] levels = new string[boxCount, 1];
            string[,] notes = new string[boxCount, 1];
            string[,] names = new string[boxCount, 1];
            string[,] descriptions = new string[boxCount, 1];
            string[,] unitPrices = new string[boxCount, 1];
            string[] links = new string[boxCount];
            string[,] dimAs = new string[boxCount, 1];
            string[,] dimBs = new string[boxCount, 1];
            string[,] dimCs = new string[boxCount, 1];
            
            int offset = 0;
            foreach (DrawerBox box in boxes) {
                lines[offset, 0] =             box.LineNumber.ToString();
                qtys[offset, 0] =              box.Qty.ToString();
                widths[offset, 0] =            box.Width.ToString();
                heights[offset, 0] =           box.Height.ToString();
                depths[offset, 0] =            box.Depth.ToString();
                materials[offset, 0] =         box.SideMaterial;
                bottoms[offset, 0] =           box.BottomMaterial;
                notches[offset, 0] =           box.NotchOption.ToString();
                inserts[offset, 0] =           box.InsertOption.ToString();
                clips[offset, 0] =             box.ClipsOption.ToString();
                mountingHoles[offset, 0] =     box.MountingHoles ? "Yes" : "No";
                finishes[offset, 0] =          box.PostFinish ? "Yes" : "No";
                scoops[offset, 0] =            box.ScoopFront ? "Yes" : "No";
                string logoSide = box.LogoInside ? "-In" : "-Out";
                logos[offset, 0] =             box.Logo ? $"Yes{logoSide}" : "No";
                levels[offset, 0] =            box.LevelName;
                notes[offset, 0] =             box.Note;
                names[offset, 0] =             box.ProductName;
                descriptions[offset, 0] =      box.ProductDescription;
                unitPrices[offset, 0] =        box.UnitPrice.ToString();
                links[offset] =             $"=HYPERLINK(\"#LineClicked()\", \"Print Label\")";

                if (box is UDrawerBox) {
                    var ubox = box as UDrawerBox;
                    dimAs[offset,0] = ubox.A.ToString();
                    dimBs[offset,0] = ubox.B.ToString();
                    dimCs[offset, 0] = ubox.C.ToString();
                } else {
                    dimAs[offset, 0] = "";
                    dimBs[offset, 0] = "";
                    dimCs[offset, 0] = "";
                }

                offset++;

            }

            outputSheet.Range[lineCol.Offset[1], lineCol.Offset[boxCount]].Value2 = lines;
            outputSheet.Range[qtyCol.Offset[1],qtyCol.Offset[boxCount]].Value2 =  qtys;
            outputSheet.Range[widthCol.Offset[1],widthCol.Offset[boxCount]].Value2 =  widths;
            outputSheet.Range[heightCol.Offset[1],heightCol.Offset[boxCount]].Value2 =  heights;
            outputSheet.Range[depthCol.Offset[1],depthCol.Offset[boxCount]].Value2 =  depths;
            outputSheet.Range[dimACol.Offset[1],dimACol.Offset[boxCount]].Value2 =  dimAs;
            outputSheet.Range[dimBCol.Offset[1],dimBCol.Offset[boxCount]].Value2 =  dimBs;
            outputSheet.Range[dimCCol.Offset[1],dimCCol.Offset[boxCount]].Value2 =  dimCs;
            outputSheet.Range[materialCol.Offset[1],materialCol.Offset[boxCount]].Value2 =  materials;
            outputSheet.Range[bottomCol.Offset[1],bottomCol.Offset[boxCount]].Value2 =  bottoms;
            outputSheet.Range[insertCol.Offset[1],insertCol.Offset[boxCount]].Value2 =  inserts;
            outputSheet.Range[notchCol.Offset[1],notchCol.Offset[boxCount]].Value2 =  notches;
            outputSheet.Range[clipCol.Offset[1],clipCol.Offset[boxCount]].Value2 =  clips;
            outputSheet.Range[mountingHolesCol.Offset[1],mountingHolesCol.Offset[boxCount]].Value2 =  mountingHoles;
            outputSheet.Range[finishCol.Offset[1],finishCol.Offset[boxCount]].Value2 =  finishes;
            outputSheet.Range[scoopCol.Offset[1],scoopCol.Offset[boxCount]].Value2 =  scoops;
            outputSheet.Range[logoCol.Offset[1],logoCol.Offset[boxCount]].Value2 =  logos;
            outputSheet.Range[levelCol.Offset[1],levelCol.Offset[boxCount]].Value2 =  levels;
            outputSheet.Range[noteCol.Offset[1],noteCol.Offset[boxCount]].Value2 =  notes;
            outputSheet.Range[nameCol.Offset[1],nameCol.Offset[boxCount]].Value2 =  names;
            outputSheet.Range[descriptionCol.Offset[1],descriptionCol.Offset[boxCount]].Value2 =  descriptions;
            outputSheet.Range[unitPriceCol.Offset[1],unitPriceCol.Offset[boxCount]].Value2 =  unitPrices;

            offset = 0;
            foreach (string link in links) {
                linkCol.Offset[offset++ + 1].Formula = link;
            }

        }

    }

}
