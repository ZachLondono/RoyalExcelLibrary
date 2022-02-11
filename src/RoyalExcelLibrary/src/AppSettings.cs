using System.Collections.Generic;

namespace RoyalExcelLibrary.ExcelUI {

    public class AppSettings {

        public IDictionary<string, InvoiceEmailConfiguration> InvoicesConfigs { get; set; }

        public TrashSettings TrashSettings { get; set; }

        public ManufacturingValues ManufacturingValues { get; set; }

        public PrinterSettings PrinterSettings { get; set; }

    }

    public class TrashSettings {

        /// <summary>
        /// This defines the width of the trash can, where it meets the 
        /// </summary>
        public double CanWidth { get; set; }

        public double CanDepth { get; set; }

        /// <summary>
        /// This defines the maximum depth of the top for a single trash top
        /// </summary>
        public double SingleTopMaxDepth { get; set; }

        /// <summary>
        /// This defines the maximum depth of the top for a double trash top
        /// </summary>
        public double DoubleTopMaxDepth { get; set; }
        public double DoubleWideTopMaxDepth { get; set; }

        /// <summary>
        /// The amount of space between the two trash can holes
        /// </summary>
        public double DoubleSpaceBetween { get; set; }

        public double CutOutRadius { get; set; }

    }

    public class InvoiceEmailConfiguration {

        public string InvoiceDirectory { get; set; }

        public string From { get; set; }

        public string[] To { get; set; }

        public string[] Cc { get; set; }

    }

    public class ManufacturingValues {

        public double DadoDepth {get; set; } = 6;
        
        public double SideAdj {get; set; } = 16;
        
        public double FrontBackAdj {get; set; } = 0.5;
        
        public double BottomAdj {get; set; } = 1;
        
        public double SideThickness {get; set; } = 16;
        
        public double SideSqrFtWeight {get; set; } = 2.1;
        
        public double BottomSqrFtWeight1_4 {get; set; } = 0.65;
        
        public double BottomSqrFtWeight1_2 {get; set; } = 1.55;

    }

    public class PrinterSettings {
        public string DefaultPrinter { get; set; } = "SHARP MX-M283N PCL6";
    }

}
