namespace SqlToExcel
{
    partial class SqlToExcel
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
            this.txtSql = new System.Windows.Forms.TextBox();
            this.btnExec = new System.Windows.Forms.Button();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnSelectPath = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.ckbIsIncludeHeader = new System.Windows.Forms.CheckBox();
            this.ckbSheet = new System.Windows.Forms.CheckBox();
            this.ckbExcel = new System.Windows.Forms.CheckBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.txtExcelName = new System.Windows.Forms.TextBox();
            this.lblExcelName = new System.Windows.Forms.Label();
            this.trvConnection = new System.Windows.Forms.TreeView();
            this.SuspendLayout();
            // 
            // txtSql
            // 
            this.txtSql.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtSql.Location = new System.Drawing.Point(14, 16);
            this.txtSql.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtSql.Multiline = true;
            this.txtSql.Name = "txtSql";
            this.txtSql.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtSql.Size = new System.Drawing.Size(458, 350);
            this.txtSql.TabIndex = 0;
            // 
            // btnExec
            // 
            this.btnExec.Location = new System.Drawing.Point(698, 406);
            this.btnExec.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnExec.Name = "btnExec";
            this.btnExec.Size = new System.Drawing.Size(87, 33);
            this.btnExec.TabIndex = 4;
            this.btnExec.Text = "执行";
            this.btnExec.UseVisualStyleBackColor = true;
            this.btnExec.Click += new System.EventHandler(this.btnExec_Click);
            // 
            // txtFilePath
            // 
            this.txtFilePath.Location = new System.Drawing.Point(14, 405);
            this.txtFilePath.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.ReadOnly = true;
            this.txtFilePath.Size = new System.Drawing.Size(363, 27);
            this.txtFilePath.TabIndex = 5;
            // 
            // btnSelectPath
            // 
            this.btnSelectPath.Location = new System.Drawing.Point(385, 405);
            this.btnSelectPath.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.btnSelectPath.Name = "btnSelectPath";
            this.btnSelectPath.Size = new System.Drawing.Size(87, 33);
            this.btnSelectPath.TabIndex = 6;
            this.btnSelectPath.Text = "导出到";
            this.btnSelectPath.UseVisualStyleBackColor = true;
            this.btnSelectPath.Click += new System.EventHandler(this.button1_Click);
            // 
            // ckbIsIncludeHeader
            // 
            this.ckbIsIncludeHeader.AutoSize = true;
            this.ckbIsIncludeHeader.Location = new System.Drawing.Point(14, 374);
            this.ckbIsIncludeHeader.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.ckbIsIncludeHeader.Name = "ckbIsIncludeHeader";
            this.ckbIsIncludeHeader.Size = new System.Drawing.Size(91, 24);
            this.ckbIsIncludeHeader.TabIndex = 7;
            this.ckbIsIncludeHeader.Text = "包含表头";
            this.ckbIsIncludeHeader.UseVisualStyleBackColor = true;
            // 
            // ckbSheet
            // 
            this.ckbSheet.AutoSize = true;
            this.ckbSheet.Location = new System.Drawing.Point(105, 374);
            this.ckbSheet.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.ckbSheet.Name = "ckbSheet";
            this.ckbSheet.Size = new System.Drawing.Size(103, 24);
            this.ckbSheet.TabIndex = 8;
            this.ckbSheet.Text = "Sheet分开";
            this.ckbSheet.UseVisualStyleBackColor = true;
            // 
            // ckbExcel
            // 
            this.ckbExcel.AutoSize = true;
            this.ckbExcel.Location = new System.Drawing.Point(203, 374);
            this.ckbExcel.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.ckbExcel.Name = "ckbExcel";
            this.ckbExcel.Size = new System.Drawing.Size(98, 24);
            this.ckbExcel.TabIndex = 9;
            this.ckbExcel.Text = "Excel分开";
            this.ckbExcel.UseVisualStyleBackColor = true;
            this.ckbExcel.CheckedChanged += new System.EventHandler(this.ckbExcel_CheckedChanged);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // txtExcelName
            // 
            this.txtExcelName.Location = new System.Drawing.Point(564, 410);
            this.txtExcelName.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.txtExcelName.Name = "txtExcelName";
            this.txtExcelName.Size = new System.Drawing.Size(116, 27);
            this.txtExcelName.TabIndex = 10;
            // 
            // lblExcelName
            // 
            this.lblExcelName.AutoSize = true;
            this.lblExcelName.Location = new System.Drawing.Point(483, 414);
            this.lblExcelName.Name = "lblExcelName";
            this.lblExcelName.Size = new System.Drawing.Size(80, 20);
            this.lblExcelName.TabIndex = 11;
            this.lblExcelName.Text = "Excel名称:";
            // 
            // trvConnection
            // 
            this.trvConnection.CheckBoxes = true;
            this.trvConnection.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel, ((byte)(134)));
            this.trvConnection.Location = new System.Drawing.Point(486, 17);
            this.trvConnection.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.trvConnection.Name = "trvConnection";
            this.trvConnection.Size = new System.Drawing.Size(299, 350);
            this.trvConnection.TabIndex = 12;
            this.trvConnection.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.trvConnection_AfterCheck);
            // 
            // SqlToExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(797, 474);
            this.Controls.Add(this.trvConnection);
            this.Controls.Add(this.lblExcelName);
            this.Controls.Add(this.txtExcelName);
            this.Controls.Add(this.ckbExcel);
            this.Controls.Add(this.ckbSheet);
            this.Controls.Add(this.ckbIsIncludeHeader);
            this.Controls.Add(this.btnSelectPath);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.btnExec);
            this.Controls.Add(this.txtSql);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "SqlToExcel";
            this.Text = "SqlToExcel";
            this.Load += new System.EventHandler(this.SqlToExcel_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtSql;
        private System.Windows.Forms.Button btnExec;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button btnSelectPath;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.CheckBox ckbIsIncludeHeader;
        private System.Windows.Forms.CheckBox ckbSheet;
        private System.Windows.Forms.CheckBox ckbExcel;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox txtExcelName;
        private System.Windows.Forms.Label lblExcelName;
        private System.Windows.Forms.TreeView trvConnection;
    }
}