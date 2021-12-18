namespace RoyalExcelLibrary.ExcelUI.Views {
    partial class VendorSelector {
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
            this.VendorList = new System.Windows.Forms.ListBox();
            this.SubmitBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // VendorList
            // 
            this.VendorList.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.VendorList.FormattingEnabled = true;
            this.VendorList.ItemHeight = 24;
            this.VendorList.Items.AddRange(new object[] {
            "Richelieu",
            "Hafele",
            "On Track",
            "Metro"});
            this.VendorList.Location = new System.Drawing.Point(12, 12);
            this.VendorList.Name = "VendorList";
            this.VendorList.Size = new System.Drawing.Size(213, 124);
            this.VendorList.TabIndex = 0;
            // 
            // SubmitBtn
            // 
            this.SubmitBtn.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.SubmitBtn.Location = new System.Drawing.Point(12, 142);
            this.SubmitBtn.Name = "SubmitBtn";
            this.SubmitBtn.Size = new System.Drawing.Size(212, 23);
            this.SubmitBtn.TabIndex = 1;
            this.SubmitBtn.Text = "Submit";
            this.SubmitBtn.UseVisualStyleBackColor = true;
            // 
            // VendorSelector
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(236, 175);
            this.Controls.Add(this.SubmitBtn);
            this.Controls.Add(this.VendorList);
            this.Name = "VendorSelector";
            this.Text = "VendorSelector";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox VendorList;
        private System.Windows.Forms.Button SubmitBtn;
    }
}