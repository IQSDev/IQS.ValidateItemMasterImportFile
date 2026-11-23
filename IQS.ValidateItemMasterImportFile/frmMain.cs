using Validate_Master_Import;

namespace IQS.ValidateItemMasterImportFile
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            btnValidate.Enabled = false;
            lblValidatedSpreadsheet.Text = string.Empty;
            lblOrigSpreadsheet.Text = string.Empty;
            txtErrors.Text = string.Empty;
            txtErrors.Visible = false;
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            txtErrors.Text = string.Empty;

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx";
                openFileDialog.Title = "Select an Excel File";
                openFileDialog.InitialDirectory = @"C:\Users\TedSturr\IQS LLC\Global Operations - General\Inventory Management System\Recently Added Items";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedFilePath = openFileDialog.FileName;
                    string newFilePath = selectedFilePath.Replace("[", "_").Replace("]", "_");
                    #region Rename File so PowerShell won't have an issue:
                    File.Move(selectedFilePath, newFilePath);
                    #endregion
                    // Display the selected file path (you can update a label or any other UI element)
                    lblOrigSpreadsheet.Text = newFilePath;
                    btnValidate.Enabled = true;
                }
            }
        }

        private void btnValidate_Click(object sender, EventArgs e)
        {
            int exitCode = 0;
            exitCode = ItemMasterValidator.Run(lblOrigSpreadsheet.Text);
            if (exitCode == 0)
            {
                lblValidatedSpreadsheet.Text = lblOrigSpreadsheet.Text.Replace(".xlsx", "_Validated.xlsx");
                txtErrors.Text = "No errors found.";
            }
            else
            {
                lblValidatedSpreadsheet.Text = "Validation failed.";
                txtErrors.Text = "See message box for details.";
            }
        }
    }
}
