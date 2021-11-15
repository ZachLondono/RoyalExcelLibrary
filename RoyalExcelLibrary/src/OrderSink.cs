using System.Linq;

using Microsoft.Office.Interop.Excel;
using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;

namespace RoyalExcelLibrary {
    public class OrderSink {

        public static void WriteToSheet(Worksheet outputSheet, Order order) {

            var customer = order.Customer;
            outputSheet.Range["CustomerName"].Value2 =      customer.Name;
            outputSheet.Range["CustomerAddress1"].Value2 =  customer.Address.Line1;
            outputSheet.Range["CustomerAddress2"].Value2 =  customer.Address.Line2;
            outputSheet.Range["CustomerCity"].Value2 =      customer.Address.City;
            outputSheet.Range["CustomerState"].Value2 =     customer.Address.State;
            outputSheet.Range["CustomerZip"].Value2 =       customer.Address.Zip;

            var vendor = order.Vendor;
            outputSheet.Range["VendorName"].Value2 =        vendor.Name;
            outputSheet.Range["VendorAddress1"].Value2 =    vendor.Address.Line1;
            outputSheet.Range["VendorAddress2"].Value2 =    vendor.Address.Line2;
            outputSheet.Range["VendorCity"].Value2 =        vendor.Address.City;
            outputSheet.Range["VendorState"].Value2 =       vendor.Address.State;
            outputSheet.Range["VendorZip"].Value2 =         vendor.Address.Zip;

            var supplier = order.Supplier;
            outputSheet.Range["SupplierName"].Value2 =      supplier.Name;
            outputSheet.Range["SupplierAddress1"].Value2 =  supplier.Address.Line1;
            outputSheet.Range["SupplierAddress2"].Value2 =  supplier.Address.Line2;
            outputSheet.Range["SupplierCity"].Value2 =      supplier.Address.City;
            outputSheet.Range["SupplierState"].Value2 =     supplier.Address.State;
            outputSheet.Range["SupplierZip"].Value2 =       supplier.Address.Zip;

            var orderNum =          outputSheet.Range["OrderNumber"];
            orderNum.Value2 =       order.Number.ToString();
            var orderSource =       outputSheet.Range["OrderSource"];
            orderSource.Value2 =    order.Job.JobSource.ToString();

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

            } else {

                orderField_1.Value2 = "Job Name";
                orderFieldValue_1.Value2 = order.Job.Name;

            }

            if (order.Job.JobSource.ToLower().Equals("allmoxy")) {
                outputSheet.Range["OrderSourceLink"].Value2 = $"https://metrodrawerboxes.allmoxy.com/orders/quote/{order.Number}/";
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
            var notchCol =      outputSheet.Range["NotchCol"];
            var clipCol =       outputSheet.Range["ClipCol"];
            var mountingHolesCol = outputSheet.Range["MountingHolesCol"];
            var finishCol =     outputSheet.Range["FinishCol"];
            var scoopCol =      outputSheet.Range["ScoopCol"];
            var levelCol =      outputSheet.Range["LevelCol"];
            var noteCol =       outputSheet.Range["NoteCol"];
            var nameCol =       outputSheet.Range["NameCol"];
            var descriptionCol =outputSheet.Range["DescriptionCol"];
            var unitPriceCol =  outputSheet.Range["UnitPriceCol"];
            var linkCol =       outputSheet.Range["LinkCol"];

            var boxes = order.Products.Where(p => p is DrawerBox).Cast<DrawerBox>();

            outputSheet.Range["BoxCount"].Value2 = boxes.Sum(b => b.Qty);

            int offset = 1;
            foreach (DrawerBox box in boxes) {

                lineCol.Offset[offset, 0].Value2 =          box.LineNumber.ToString();
                qtyCol.Offset[offset, 0].Value2 =           box.Qty.ToString();
                widthCol.Offset[offset, 0].Value2 =         box.Width.ToString();
                heightCol.Offset[offset, 0].Value2 =        box.Height.ToString();
                depthCol.Offset[offset, 0].Value2 =         box.Depth.ToString();
                materialCol.Offset[offset, 0].Value2 =      box.SideMaterial.ToString();
                bottomCol.Offset[offset, 0].Value2 =        box.BottomMaterial.ToString();
                notchCol.Offset[offset, 0].Value2 =         box.NotchOption.ToString();
                clipCol.Offset[offset, 0].Value2 =          box.ClipsOption.ToString();
                mountingHolesCol.Offset[offset, 0].Value2 = box.MountingHoles ? "Yes" : "No";
                finishCol.Offset[offset, 0].Value2 =        box.PostFinish ? "Yes" : "No";
                scoopCol.Offset[offset, 0].Value2 =         box.ScoopFront ? "Yes" : "No";
                levelCol.Offset[offset, 0].Value2 =         box.LevelName;
                noteCol.Offset[offset, 0].Value2 =          box.Note;
                nameCol.Offset[offset, 0].Value2 =          box.ProductName;
                descriptionCol.Offset[offset, 0].Value2 =   box.ProductDescription;
                unitPriceCol.Offset[offset, 0].Value2 =     box.UnitPrice.ToString();
                linkCol.Offset[offset,0].Formula =          $"=HYPERLINK(\"#LineClicked({offset})\", \"Print Label\")";

                if (box is UDrawerBox) {
                    var ubox = box as UDrawerBox;
                    dimACol.Offset[offset, 0].Value2 = ubox.A.ToString();
                    dimBCol.Offset[offset, 0].Value2 = ubox.B.ToString();
                    dimCCol.Offset[offset, 0].Value2 = ubox.C.ToString();
                }

                offset++;

            }
            
        }

    }

}
