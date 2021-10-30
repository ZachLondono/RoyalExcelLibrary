
namespace RoyalExcelLibrary.Views {
	partial class ErrorMessage {
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing) {
			if (disposing && (components != null)) {
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent() {
			this.errorTitle = new System.Windows.Forms.Label();
			this.errorSummary = new System.Windows.Forms.Label();
			this.errorDetail = new System.Windows.Forms.RichTextBox();
			this.sendBtn = new System.Windows.Forms.Button();
			this.closeBtn = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// errorTitle
			// 
			this.errorTitle.AutoSize = true;
			this.errorTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.errorTitle.Location = new System.Drawing.Point(13, 13);
			this.errorTitle.Name = "errorTitle";
			this.errorTitle.Size = new System.Drawing.Size(88, 20);
			this.errorTitle.TabIndex = 0;
			this.errorTitle.Text = "Error Title";
			// 
			// errorSummary
			// 
			this.errorSummary.AutoSize = true;
			this.errorSummary.Location = new System.Drawing.Point(14, 43);
			this.errorSummary.Name = "errorSummary";
			this.errorSummary.Size = new System.Drawing.Size(97, 16);
			this.errorSummary.TabIndex = 1;
			this.errorSummary.Text = "Error Summary";
			// 
			// errorDetail
			// 
			this.errorDetail.Location = new System.Drawing.Point(12, 78);
			this.errorDetail.Name = "errorDetail";
			this.errorDetail.ReadOnly = true;
			this.errorDetail.Size = new System.Drawing.Size(498, 197);
			this.errorDetail.TabIndex = 2;
			this.errorDetail.Text = "";
			// 
			// sendBtn
			// 
			this.sendBtn.Location = new System.Drawing.Point(13, 282);
			this.sendBtn.Name = "sendBtn";
			this.sendBtn.Size = new System.Drawing.Size(146, 23);
			this.sendBtn.TabIndex = 3;
			this.sendBtn.Text = "Send Error Report";
			this.sendBtn.UseVisualStyleBackColor = true;
			// 
			// closeBtn
			// 
			this.closeBtn.Location = new System.Drawing.Point(434, 282);
			this.closeBtn.Name = "closeBtn";
			this.closeBtn.Size = new System.Drawing.Size(75, 23);
			this.closeBtn.TabIndex = 4;
			this.closeBtn.Text = "Close";
			this.closeBtn.UseVisualStyleBackColor = true;
			this.closeBtn.Click += new System.EventHandler(this.closeBtn_Click);
			// 
			// ErrorMessage
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(522, 316);
			this.Controls.Add(this.closeBtn);
			this.Controls.Add(this.sendBtn);
			this.Controls.Add(this.errorDetail);
			this.Controls.Add(this.errorSummary);
			this.Controls.Add(this.errorTitle);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.Name = "ErrorMessage";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Encountered Error";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label errorTitle;
		private System.Windows.Forms.Label errorSummary;
		private System.Windows.Forms.RichTextBox errorDetail;
		private System.Windows.Forms.Button sendBtn;
		private System.Windows.Forms.Button closeBtn;
	}
}