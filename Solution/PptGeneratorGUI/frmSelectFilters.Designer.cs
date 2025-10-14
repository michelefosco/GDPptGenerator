namespace PptGeneratorGUI
{
    partial class frmSelectFilters
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.cblFilters = new System.Windows.Forms.CheckedListBox();
            this.lblFieldName = new System.Windows.Forms.Label();
            this.btnClear = new System.Windows.Forms.Button();
            this.bntSelectAll = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnOk
            // 
            this.btnOk.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnOk.Location = new System.Drawing.Point(69, 525);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(104, 23);
            this.btnOk.TabIndex = 4;
            this.btnOk.Text = "&Ok";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(179, 525);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(104, 23);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "&Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // cblFilters
            // 
            this.cblFilters.CausesValidation = false;
            this.cblFilters.CheckOnClick = true;
            this.cblFilters.FormattingEnabled = true;
            this.cblFilters.Location = new System.Drawing.Point(5, 50);
            this.cblFilters.Name = "cblFilters";
            this.cblFilters.Size = new System.Drawing.Size(346, 469);
            this.cblFilters.TabIndex = 3;
            // 
            // lblFieldName
            // 
            this.lblFieldName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFieldName.Location = new System.Drawing.Point(0, 4);
            this.lblFieldName.Name = "lblFieldName";
            this.lblFieldName.Size = new System.Drawing.Size(351, 16);
            this.lblFieldName.TabIndex = 4;
            this.lblFieldName.Text = "XXXXXXXXXXXXXXXXXXXXXXXXXXX";
            this.lblFieldName.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(295, 23);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(58, 20);
            this.btnClear.TabIndex = 2;
            this.btnClear.Text = "&Clear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // bntSelectAll
            // 
            this.bntSelectAll.Location = new System.Drawing.Point(231, 23);
            this.bntSelectAll.Name = "bntSelectAll";
            this.bntSelectAll.Size = new System.Drawing.Size(58, 20);
            this.bntSelectAll.TabIndex = 1;
            this.bntSelectAll.Text = "Select &all";
            this.bntSelectAll.UseVisualStyleBackColor = true;
            this.bntSelectAll.Click += new System.EventHandler(this.bntSelectAll_Click);
            // 
            // frmSelectFilters
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(357, 550);
            this.Controls.Add(this.bntSelectAll);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.lblFieldName);
            this.Controls.Add(this.cblFilters);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmSelectFilters";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Select filters";
            this.Load += new System.EventHandler(this.frmSelectFilters_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.CheckedListBox cblFilters;
        private System.Windows.Forms.Label lblFieldName;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Button bntSelectAll;
    }
}