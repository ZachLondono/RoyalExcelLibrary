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
	public partial class ErrorMessage : Form {
		
		public ErrorMessage() {
			InitializeComponent();
			sendBtn.Hide();
		}

		public void SetError(string title, string summary, string detailed) {
			errorTitle.Text = title;
			errorSummary.Text = summary;
			errorDetail.Text = detailed;
		}

		private void closeBtn_Click(object sender, EventArgs e) {
			Close();
		}

	}
}
