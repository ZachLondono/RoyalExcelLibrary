using Microsoft.Office.Interop.Excel;

namespace RoyalExcelLibrary.Providers {
    internal interface IExcelOrderProvider : IOrderProvider {

        Application App { get; set; }

    }

    internal interface IFileOrderProvider : IOrderProvider {

        string FilePath { get; set; }

    }

}
