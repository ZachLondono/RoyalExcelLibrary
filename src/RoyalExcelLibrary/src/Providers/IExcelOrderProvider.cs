using Excel = Microsoft.Office.Interop.Excel;

namespace RoyalExcelLibrary.ExcelUI.Providers {
    internal interface IExcelOrderProvider : IOrderProvider {

        Excel.Application App { get; set; }

    }

    internal interface IFileOrderProvider : IOrderProvider {

        string FilePath { get; set; }

    }

}
