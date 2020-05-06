namespace ReportMaster
{
    partial class MainForm
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnCreateXLSFileReport = new System.Windows.Forms.Button();
            this.tbCSVFileName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCSVOpen = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnCreateXLSFileReport
            // 
            this.btnCreateXLSFileReport.Location = new System.Drawing.Point(187, 89);
            this.btnCreateXLSFileReport.Name = "btnCreateXLSFileReport";
            this.btnCreateXLSFileReport.Size = new System.Drawing.Size(75, 31);
            this.btnCreateXLSFileReport.TabIndex = 0;
            this.btnCreateXLSFileReport.Text = "Run";
            this.btnCreateXLSFileReport.UseVisualStyleBackColor = true;
            this.btnCreateXLSFileReport.Click += new System.EventHandler(this.btnOpenXLSFile_Click);
            // 
            // tbCSVFileName
            // 
            this.tbCSVFileName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbCSVFileName.Location = new System.Drawing.Point(72, 35);
            this.tbCSVFileName.Name = "tbCSVFileName";
            this.tbCSVFileName.Size = new System.Drawing.Size(259, 21);
            this.tbCSVFileName.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(69, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "CSV-file:";
            // 
            // btnCSVOpen
            // 
            this.btnCSVOpen.Location = new System.Drawing.Point(348, 33);
            this.btnCSVOpen.Name = "btnCSVOpen";
            this.btnCSVOpen.Size = new System.Drawing.Size(27, 23);
            this.btnCSVOpen.TabIndex = 3;
            this.btnCSVOpen.Text = "...";
            this.btnCSVOpen.UseVisualStyleBackColor = true;
            this.btnCSVOpen.Click += new System.EventHandler(this.btnCSVOpen_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(448, 159);
            this.Controls.Add(this.btnCSVOpen);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbCSVFileName);
            this.Controls.Add(this.btnCreateXLSFileReport);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Мастер отчетов CSV to XLSX";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCreateXLSFileReport;
        private System.Windows.Forms.TextBox tbCSVFileName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCSVOpen;
    }
}

