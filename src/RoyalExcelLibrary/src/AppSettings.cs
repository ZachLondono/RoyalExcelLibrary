using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoyalExcelLibrary.ExcelUI {

    public class AppSettings {

        public TrashSettings TrashSettings { get; set; }

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


}
