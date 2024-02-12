using System.Windows.Forms;

namespace ExcelHierarchyConversion_InterOp
{
    partial class ExcelHierarchyCon
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
            this.fileDialog = new System.Windows.Forms.OpenFileDialog();
            this.uploadButton = new System.Windows.Forms.Button();
            this.outputButton = new System.Windows.Forms.Button();
            this.convertButton = new System.Windows.Forms.Button();
            this.folderBrowse = new System.Windows.Forms.FolderBrowserDialog();
            this.outputPathTextBox = new System.Windows.Forms.TextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.exitButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.inputPathTextBox = new System.Windows.Forms.TextBox();
            this.verificationPathTextBox = new System.Windows.Forms.TextBox();
            this.uploadVerificationButton = new System.Windows.Forms.Button();
            this.checkBox_LogErrors = new System.Windows.Forms.CheckBox();
            this.CheckBox_splitFiles = new System.Windows.Forms.CheckBox();
            this.txtBox_inputPathMaximo = new System.Windows.Forms.TextBox();
            this.btn_UploadMaximoSheet = new System.Windows.Forms.Button();
            this.label_operationStatus = new System.Windows.Forms.Label();
            this.label_fixed = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // fileDialog
            // 
            this.fileDialog.FileName = "openFileDialog1";
            // 
            // uploadButton
            // 
            this.uploadButton.BackColor = System.Drawing.Color.White;
            this.uploadButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uploadButton.Location = new System.Drawing.Point(15, 110);
            this.uploadButton.Name = "uploadButton";
            this.uploadButton.Size = new System.Drawing.Size(130, 31);
            this.uploadButton.TabIndex = 2;
            this.uploadButton.Text = "Upload Input File";
            this.uploadButton.UseVisualStyleBackColor = false;
            this.uploadButton.Click += new System.EventHandler(this.uploadButton_Click);
            // 
            // outputButton
            // 
            this.outputButton.BackColor = System.Drawing.Color.White;
            this.outputButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.outputButton.Location = new System.Drawing.Point(15, 185);
            this.outputButton.Name = "outputButton";
            this.outputButton.Size = new System.Drawing.Size(130, 31);
            this.outputButton.TabIndex = 3;
            this.outputButton.Text = "Output Directory";
            this.outputButton.UseVisualStyleBackColor = false;
            this.outputButton.Click += new System.EventHandler(this.button2_Click);
            // 
            // convertButton
            // 
            this.convertButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.convertButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.convertButton.Location = new System.Drawing.Point(12, 275);
            this.convertButton.Name = "convertButton";
            this.convertButton.Size = new System.Drawing.Size(130, 52);
            this.convertButton.TabIndex = 4;
            this.convertButton.Text = "Convert";
            this.convertButton.UseVisualStyleBackColor = false;
            this.convertButton.Click += new System.EventHandler(this.convertButton_Click);
            // 
            // outputPathTextBox
            // 
            this.outputPathTextBox.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.outputPathTextBox.Location = new System.Drawing.Point(151, 185);
            this.outputPathTextBox.Multiline = true;
            this.outputPathTextBox.Name = "outputPathTextBox";
            this.outputPathTextBox.Size = new System.Drawing.Size(412, 31);
            this.outputPathTextBox.TabIndex = 5;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 259);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(130, 10);
            this.progressBar1.TabIndex = 6;
            this.progressBar1.Visible = false;
            // 
            // exitButton
            // 
            this.exitButton.BackColor = System.Drawing.Color.Red;
            this.exitButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.exitButton.Location = new System.Drawing.Point(571, 346);
            this.exitButton.Name = "exitButton";
            this.exitButton.Size = new System.Drawing.Size(66, 43);
            this.exitButton.TabIndex = 8;
            this.exitButton.Text = "Exit";
            this.exitButton.UseVisualStyleBackColor = false;
            this.exitButton.Click += new System.EventHandler(this.exitButton_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft YaHei UI", 15.75F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
                | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(41, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(432, 98);
            this.label1.TabIndex = 9;
            this.label1.Text = "Excel hierarchy Converter";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // inputPathTextBox
            // 
            this.inputPathTextBox.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.inputPathTextBox.Location = new System.Drawing.Point(151, 110);
            this.inputPathTextBox.Multiline = true;
            this.inputPathTextBox.Name = "inputPathTextBox";
            this.inputPathTextBox.Size = new System.Drawing.Size(412, 31);
            this.inputPathTextBox.TabIndex = 10;
            this.inputPathTextBox.TextChanged += new System.EventHandler(this.inputPathTextBox_TextChanged);
            // 
            // verificationPathTextBox
            // 
            this.verificationPathTextBox.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.verificationPathTextBox.Location = new System.Drawing.Point(151, 147);
            this.verificationPathTextBox.Multiline = true;
            this.verificationPathTextBox.Name = "verificationPathTextBox";
            this.verificationPathTextBox.Size = new System.Drawing.Size(412, 32);
            this.verificationPathTextBox.TabIndex = 11;
            this.verificationPathTextBox.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // uploadVerificationButton
            // 
            this.uploadVerificationButton.BackColor = System.Drawing.Color.White;
            this.uploadVerificationButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uploadVerificationButton.Location = new System.Drawing.Point(15, 147);
            this.uploadVerificationButton.Name = "uploadVerificationButton";
            this.uploadVerificationButton.Size = new System.Drawing.Size(130, 32);
            this.uploadVerificationButton.TabIndex = 12;
            this.uploadVerificationButton.Text = "Upload Verification File";
            this.uploadVerificationButton.UseVisualStyleBackColor = false;
            this.uploadVerificationButton.Click += new System.EventHandler(this.uploadVerificationButton_Click);
            // 
            // checkBox_LogErrors
            // 
            this.checkBox_LogErrors.AutoSize = true;
            this.checkBox_LogErrors.BackColor = System.Drawing.Color.LightGray;
            this.checkBox_LogErrors.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkBox_LogErrors.Location = new System.Drawing.Point(151, 259);
            this.checkBox_LogErrors.Name = "checkBox_LogErrors";
            this.checkBox_LogErrors.Size = new System.Drawing.Size(171, 17);
            this.checkBox_LogErrors.TabIndex = 13;
            this.checkBox_LogErrors.Text = "Create Log File For Errors";
            this.checkBox_LogErrors.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.checkBox_LogErrors.UseVisualStyleBackColor = false;
            // 
            // CheckBox_splitFiles
            // 
            this.CheckBox_splitFiles.AutoSize = true;
            this.CheckBox_splitFiles.BackColor = System.Drawing.Color.LightGray;
            this.CheckBox_splitFiles.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.CheckBox_splitFiles.Location = new System.Drawing.Point(151, 282);
            this.CheckBox_splitFiles.Name = "CheckBox_splitFiles";
            this.CheckBox_splitFiles.Size = new System.Drawing.Size(81, 17);
            this.CheckBox_splitFiles.TabIndex = 14;
            this.CheckBox_splitFiles.Text = "Split Files";
            this.CheckBox_splitFiles.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.CheckBox_splitFiles.UseVisualStyleBackColor = false;
            // 
            // txtBox_inputPathMaximo
            // 
            this.txtBox_inputPathMaximo.BackColor = System.Drawing.SystemColors.InactiveBorder;
            this.txtBox_inputPathMaximo.Location = new System.Drawing.Point(151, 223);
            this.txtBox_inputPathMaximo.Multiline = true;
            this.txtBox_inputPathMaximo.Name = "txtBox_inputPathMaximo";
            this.txtBox_inputPathMaximo.Size = new System.Drawing.Size(412, 30);
            this.txtBox_inputPathMaximo.TabIndex = 16;
            // 
            // btn_UploadMaximoSheet
            // 
            this.btn_UploadMaximoSheet.BackColor = System.Drawing.Color.White;
            this.btn_UploadMaximoSheet.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_UploadMaximoSheet.Location = new System.Drawing.Point(15, 222);
            this.btn_UploadMaximoSheet.Name = "btn_UploadMaximoSheet";
            this.btn_UploadMaximoSheet.Size = new System.Drawing.Size(130, 31);
            this.btn_UploadMaximoSheet.TabIndex = 17;
            this.btn_UploadMaximoSheet.Text = "Upload Maximo Sheet";
            this.btn_UploadMaximoSheet.UseVisualStyleBackColor = false;
            this.btn_UploadMaximoSheet.Click += new System.EventHandler(this.btn_UploadMaximoSheet_Click);
            // 
            // label_operationStatus
            // 
            this.label_operationStatus.AutoSize = true;
            this.label_operationStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_operationStatus.Location = new System.Drawing.Point(104, 361);
            this.label_operationStatus.Name = "label_operationStatus";
            this.label_operationStatus.Size = new System.Drawing.Size(41, 13);
            this.label_operationStatus.TabIndex = 18;
            this.label_operationStatus.Text = "label2";
            this.label_operationStatus.Visible = false;
            // 
            // label_fixed
            // 
            this.label_fixed.AutoSize = true;
            this.label_fixed.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label_fixed.Location = new System.Drawing.Point(28, 361);
            this.label_fixed.Name = "label_fixed";
            this.label_fixed.Size = new System.Drawing.Size(58, 13);
            this.label_fixed.TabIndex = 19;
            this.label_fixed.Text = "Status ->";
            this.label_fixed.Visible = false;
            // 
            // ExcelHierarchyCon
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightGray;
            this.ClientSize = new System.Drawing.Size(635, 387);
            this.Controls.Add(this.label_fixed);
            this.Controls.Add(this.label_operationStatus);
            this.Controls.Add(this.btn_UploadMaximoSheet);
            this.Controls.Add(this.txtBox_inputPathMaximo);
            this.Controls.Add(this.CheckBox_splitFiles);
            this.Controls.Add(this.checkBox_LogErrors);
            this.Controls.Add(this.uploadVerificationButton);
            this.Controls.Add(this.verificationPathTextBox);
            this.Controls.Add(this.inputPathTextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.exitButton);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.outputPathTextBox);
            this.Controls.Add(this.convertButton);
            this.Controls.Add(this.outputButton);
            this.Controls.Add(this.uploadButton);
            this.Name = "ExcelHierarchyCon";
            this.Text = "Excel Hierarchy Converter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.OpenFileDialog fileDialog;
        private System.Windows.Forms.Button uploadButton;
        private System.Windows.Forms.Button outputButton;
        private System.Windows.Forms.Button convertButton;
        private System.Windows.Forms.FolderBrowserDialog folderBrowse;
        private System.Windows.Forms.TextBox outputPathTextBox;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button exitButton;
        private Label label1;
        private System.Windows.Forms.TextBox inputPathTextBox;
        private System.Windows.Forms.TextBox verificationPathTextBox;
        private System.Windows.Forms.Button uploadVerificationButton;
        private System.Windows.Forms.CheckBox checkBox_LogErrors;
        private System.Windows.Forms.CheckBox CheckBox_splitFiles;
        private System.Windows.Forms.TextBox txtBox_inputPathMaximo;
        private System.Windows.Forms.Button btn_UploadMaximoSheet;
        private Label label_operationStatus;
        private Label label_fixed;
    }
}

