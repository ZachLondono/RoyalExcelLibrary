using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RoyalExcelLibrary.ExcelUI.Views {
    public partial class VendorSelector : Form {
        public VendorSelector() {
            InitializeComponent();
        }

        public string GetSelected() {

            if (string.IsNullOrEmpty((string) VendorList.SelectedItem))
                return string.Empty;

            return (string)VendorList.SelectedItem;

        }

    }
}
