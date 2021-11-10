using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using RoyalExcelLibrary.Models;
using RoyalExcelLibrary.Models.Products;

namespace RoyalExcelLibrary {
    public class OrderSink {

        public static void WriteToSheet(Worksheet outputSheet, Order order) {

            var custName =      outputSheet.Range["CustomerName"];
            var custAddress1 =  outputSheet.Range["CustomerAddress1"];
            var custAddress2 =  outputSheet.Range["CustomerAddress2"];
            var custCity =      outputSheet.Range["CustomerCity"];
            var custState =     outputSheet.Range["CustomerState"];
            var custZip =       outputSheet.Range["CustomerZip"];

            var vendorName =     outputSheet.Range["VendorName"];
            var vendorAddress1 = outputSheet.Range["VendorAddress1"];
            var vendorAddress2 = outputSheet.Range["VendorAddress2"];
            var vendorCity =     outputSheet.Range["VendorCity"];
            var vendorState =    outputSheet.Range["VendorState"];
            var vendorZip =      outputSheet.Range["VendorZip"];

            var supplierName =     outputSheet.Range["SupplierName"];
            var supplierAddress1 = outputSheet.Range["SupplierAddress1"];
            var supplierAddress2 = outputSheet.Range["SupplierAddress2"];
            var supplierCity =     outputSheet.Range["SupplierCity"];
            var supplierState =    outputSheet.Range["SupplierState"];
            var supplierZip =      outputSheet.Range["SupplierZip"];

            var orderNum = outputSheet.Range["OrderNumber"];
            orderNum.Value2 = order.Number.ToString();
            var orderSource = outputSheet.Range["OrderSource"];
            orderSource.Value2 = order.Job.JobSource.ToString();

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

            var subTotal = outputSheet.Range["SubTotal"];
            subTotal.Value2 = order.SubTotal.ToString();
            var tax = outputSheet.Range["Tax"];
            tax.Value2 = order.Tax.ToString();
            var shipping = outputSheet.Range["ShippingCost"];
            shipping.Value2 = order.ShippingCost.ToString();
            var total = outputSheet.Range["TotalCost"];
            total.Value2 = (order.SubTotal + order.Tax + order.ShippingCost).ToString();


            var lineCol = outputSheet.Range["LineCol"];
            var qtyCol = outputSheet.Range["QtyCol"];
            var widthCol = outputSheet.Range["WidthCol"];
            var heightCol = outputSheet.Range["HeightCol"];
            var depthCol = outputSheet.Range["DepthCol"];
            var materialCol = outputSheet.Range["MaterialCol"];
            var bottomCol = outputSheet.Range["BottomCol"];
            var notchCol = outputSheet.Range["NotchCol"];
            var clipCol = outputSheet.Range["ClipCol"];
            var mountingHolesCol = outputSheet.Range["MountingHolesCol"];
            var finishCol = outputSheet.Range["FinishCol"];
            var scoopCol = outputSheet.Range["ScoopCol"];
            var levelCol = outputSheet.Range["LevelCol"];
            var noteCol = outputSheet.Range["NoteCol"];
            var unitPriceCol = outputSheet.Range["UnitPriceCol"];
            var linkCol = outputSheet.Range["LinkCol"];

            var boxes = order.Products.Where(p => p is DrawerBox).Cast<DrawerBox>();

            int offset = 1;
            foreach (DrawerBox box in boxes) {

                lineCol.Offset[offset, 0].Value2 = box.LineNumber.ToString();
                qtyCol.Offset[offset, 0].Value2 = box.Qty.ToString();
                widthCol.Offset[offset, 0].Value2 = box.Width.ToString();
                heightCol.Offset[offset, 0].Value2 = box.Height.ToString();
                depthCol.Offset[offset, 0].Value2 = box.Depth.ToString();
                materialCol.Offset[offset, 0].Value2 = box.SideMaterial.ToString();
                bottomCol.Offset[offset, 0].Value2 = box.BottomMaterial.ToString();
                notchCol.Offset[offset, 0].Value2 = box.NotchOption.ToString();
                clipCol.Offset[offset, 0].Value2 = box.ClipsOption.ToString();
                mountingHolesCol.Offset[offset, 0].Value2 = box.MountingHoles ? "Yes" : "No";
                finishCol.Offset[offset, 0].Value2 = box.PostFinish ? "Yes" : "No";
                scoopCol.Offset[offset, 0].Value2 = box.ScoopFront ? "Yes" : "No";
                levelCol.Offset[offset, 0].Value2 = "";
                noteCol.Offset[offset, 0].Value2 = "";
                unitPriceCol.Offset[offset, 0].Value2 = box.UnitPrice.ToString();
                linkCol.Offset[offset,0].Formula = $"HYPERLINK(\"#LineClicked({offset})\", Print Label)";

            }
            
        }

    }

}
