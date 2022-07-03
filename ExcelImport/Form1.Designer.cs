namespace ExcelImport
{
    partial class Form1
    {
        /// <summary>
        /// Wymagana zmienna projektanta.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Wyczyść wszystkie używane zasoby.
        /// </summary>
        /// <param name="disposing">prawda, jeżeli zarządzane zasoby powinny zostać zlikwidowane; Fałsz w przeciwnym wypadku.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Kod generowany przez Projektanta formularzy systemu Windows

        /// <summary>
        /// Metoda wymagana do obsługi projektanta — nie należy modyfikować
        /// jej zawartości w edytorze kodu.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnDataSource = new System.Windows.Forms.Button();
            this.btnSubmit = new System.Windows.Forms.Button();
            this.btnBrowse2 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtDataSource = new System.Windows.Forms.TextBox();
            this.txtExcelFilePath = new System.Windows.Forms.TextBox();
            this.txtDatabaseName = new System.Windows.Forms.TextBox();
            this.txtTableName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtCsvFilePath = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnDataSource
            // 
            this.btnDataSource.Location = new System.Drawing.Point(455, 51);
            this.btnDataSource.Margin = new System.Windows.Forms.Padding(4);
            this.btnDataSource.Name = "btnDataSource";
            this.btnDataSource.Size = new System.Drawing.Size(100, 28);
            this.btnDataSource.TabIndex = 0;
            this.btnDataSource.Text = "Browse";
            this.btnDataSource.UseVisualStyleBackColor = true;
            this.btnDataSource.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnSubmit
            // 
            this.btnSubmit.Location = new System.Drawing.Point(333, 274);
            this.btnSubmit.Margin = new System.Windows.Forms.Padding(4);
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.Size = new System.Drawing.Size(222, 28);
            this.btnSubmit.TabIndex = 1;
            this.btnSubmit.Text = "Submit";
            this.btnSubmit.UseVisualStyleBackColor = true;
            this.btnSubmit.Click += new System.EventHandler(this.btnSubmit_Click);
            // 
            // btnBrowse2
            // 
            this.btnBrowse2.Location = new System.Drawing.Point(455, 131);
            this.btnBrowse2.Margin = new System.Windows.Forms.Padding(4);
            this.btnBrowse2.Name = "btnBrowse2";
            this.btnBrowse2.Size = new System.Drawing.Size(100, 28);
            this.btnBrowse2.TabIndex = 2;
            this.btnBrowse2.Text = "Browse";
            this.btnBrowse2.UseVisualStyleBackColor = true;
            this.btnBrowse2.Click += new System.EventHandler(this.btnBrowse2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 16);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(87, 17);
            this.label1.TabIndex = 4;
            this.label1.Text = "Data Source";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(17, 57);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 17);
            this.label2.TabIndex = 5;
            this.label2.Text = "Excel File Path";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(17, 184);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(110, 17);
            this.label3.TabIndex = 6;
            this.label3.Text = "Database Name";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 226);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(85, 17);
            this.label4.TabIndex = 7;
            this.label4.Text = "Table Name";
            // 
            // txtDataSource
            // 
            this.txtDataSource.Location = new System.Drawing.Point(131, 13);
            this.txtDataSource.Margin = new System.Windows.Forms.Padding(4);
            this.txtDataSource.Name = "txtDataSource";
            this.txtDataSource.Size = new System.Drawing.Size(307, 23);
            this.txtDataSource.TabIndex = 8;
            // 
            // txtExcelFilePath
            // 
            this.txtExcelFilePath.Location = new System.Drawing.Point(131, 54);
            this.txtExcelFilePath.Margin = new System.Windows.Forms.Padding(4);
            this.txtExcelFilePath.Name = "txtExcelFilePath";
            this.txtExcelFilePath.Size = new System.Drawing.Size(307, 23);
            this.txtExcelFilePath.TabIndex = 9;
            // 
            // txtDatabaseName
            // 
            this.txtDatabaseName.Location = new System.Drawing.Point(131, 181);
            this.txtDatabaseName.Margin = new System.Windows.Forms.Padding(4);
            this.txtDatabaseName.Name = "txtDatabaseName";
            this.txtDatabaseName.Size = new System.Drawing.Size(307, 23);
            this.txtDatabaseName.TabIndex = 10;
            // 
            // txtTableName
            // 
            this.txtTableName.Location = new System.Drawing.Point(131, 223);
            this.txtTableName.Margin = new System.Windows.Forms.Padding(4);
            this.txtTableName.Name = "txtTableName";
            this.txtTableName.Size = new System.Drawing.Size(307, 23);
            this.txtTableName.TabIndex = 11;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(17, 137);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(94, 17);
            this.label5.TabIndex = 12;
            this.label5.Text = "CSV File Path";
            // 
            // txtCsvFilePath
            // 
            this.txtCsvFilePath.Location = new System.Drawing.Point(131, 134);
            this.txtCsvFilePath.Name = "txtCsvFilePath";
            this.txtCsvFilePath.Size = new System.Drawing.Size(307, 23);
            this.txtCsvFilePath.TabIndex = 13;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(17, 98);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(398, 17);
            this.label6.TabIndex = 14;
            this.label6.Text = "(Excel file will be converted into CSV and saved on hard drive)";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(592, 315);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtCsvFilePath);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtTableName);
            this.Controls.Add(this.txtDatabaseName);
            this.Controls.Add(this.txtExcelFilePath);
            this.Controls.Add(this.txtDataSource);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnBrowse2);
            this.Controls.Add(this.btnSubmit);
            this.Controls.Add(this.btnDataSource);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnDataSource;
        private System.Windows.Forms.Button btnSubmit;
        private System.Windows.Forms.Button btnBrowse2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtDataSource;
        private System.Windows.Forms.TextBox txtExcelFilePath;
        private System.Windows.Forms.TextBox txtDatabaseName;
        private System.Windows.Forms.TextBox txtTableName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtCsvFilePath;
        private System.Windows.Forms.Label label6;
    }
}

