namespace IQS.ValidateItemMasterImportFile
{
    partial class frmMain
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            label1 = new Label();
            btnSelectFile = new Button();
            btnValidate = new Button();
            label2 = new Label();
            lblOrigSpreadsheet = new Label();
            lblValidatedSpreadsheet = new Label();
            txtErrors = new TextBox();
            SuspendLayout();
            // 
            // label1
            // 
            label1.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            label1.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            label1.ImageAlign = ContentAlignment.MiddleRight;
            label1.Location = new Point(12, 9);
            label1.Name = "label1";
            label1.Size = new Size(118, 23);
            label1.TabIndex = 0;
            label1.Text = "Original Excel File";
            label1.TextAlign = ContentAlignment.MiddleRight;
            // 
            // btnSelectFile
            // 
            btnSelectFile.BackColor = Color.NavajoWhite;
            btnSelectFile.FlatAppearance.BorderColor = Color.Black;
            btnSelectFile.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            btnSelectFile.Location = new Point(1081, 10);
            btnSelectFile.Name = "btnSelectFile";
            btnSelectFile.Size = new Size(75, 23);
            btnSelectFile.TabIndex = 2;
            btnSelectFile.TabStop = false;
            btnSelectFile.Text = "Select File";
            btnSelectFile.UseVisualStyleBackColor = false;
            btnSelectFile.Click += btnSelectFile_Click;
            // 
            // btnValidate
            // 
            btnValidate.BackColor = Color.NavajoWhite;
            btnValidate.FlatAppearance.BorderColor = Color.Black;
            btnValidate.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            btnValidate.Location = new Point(523, 52);
            btnValidate.Name = "btnValidate";
            btnValidate.Size = new Size(189, 45);
            btnValidate.TabIndex = 3;
            btnValidate.TabStop = false;
            btnValidate.Text = "Validate Excel File";
            btnValidate.UseVisualStyleBackColor = false;
            btnValidate.Click += btnValidate_Click;
            // 
            // label2
            // 
            label2.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            label2.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            label2.ImageAlign = ContentAlignment.MiddleRight;
            label2.Location = new Point(12, 123);
            label2.Name = "label2";
            label2.Size = new Size(118, 23);
            label2.TabIndex = 4;
            label2.Text = "Validated Excel File";
            label2.TextAlign = ContentAlignment.MiddleRight;
            // 
            // lblOrigSpreadsheet
            // 
            lblOrigSpreadsheet.Location = new Point(136, 9);
            lblOrigSpreadsheet.Name = "lblOrigSpreadsheet";
            lblOrigSpreadsheet.Size = new Size(939, 23);
            lblOrigSpreadsheet.TabIndex = 6;
            lblOrigSpreadsheet.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // lblValidatedSpreadsheet
            // 
            lblValidatedSpreadsheet.Location = new Point(136, 123);
            lblValidatedSpreadsheet.Name = "lblValidatedSpreadsheet";
            lblValidatedSpreadsheet.Size = new Size(939, 23);
            lblValidatedSpreadsheet.TabIndex = 7;
            lblValidatedSpreadsheet.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // txtErrors
            // 
            txtErrors.BackColor = SystemColors.Info;
            txtErrors.Location = new Point(12, 35);
            txtErrors.Multiline = true;
            txtErrors.Name = "txtErrors";
            txtErrors.ScrollBars = ScrollBars.Horizontal;
            txtErrors.Size = new Size(505, 85);
            txtErrors.TabIndex = 8;
            txtErrors.Visible = false;
            // 
            // frmMain
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.Info;
            ClientSize = new Size(1168, 159);
            Controls.Add(txtErrors);
            Controls.Add(lblValidatedSpreadsheet);
            Controls.Add(lblOrigSpreadsheet);
            Controls.Add(label2);
            Controls.Add(btnValidate);
            Controls.Add(btnSelectFile);
            Controls.Add(label1);
            Name = "frmMain";
            Text = "Validate Item Master Import File";
            Load += frmMain_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private Button btnSelectFile;
        private Button btnValidate;
        private Label label2;
        private Label lblOrigSpreadsheet;
        private Label lblValidatedSpreadsheet;
        private TextBox txtErrors;
    }
}
