namespace zsbApp.OfficeToPdf.Test
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnOpenFolder = new System.Windows.Forms.Button();
            this.txbOfficeFile = new System.Windows.Forms.TextBox();
            this.txbPdf = new System.Windows.Forms.TextBox();
            this.btnSelectOfficeFile = new System.Windows.Forms.Button();
            this.btnToPdf = new System.Windows.Forms.Button();
            this.txbInterface = new System.Windows.Forms.TextBox();
            this.btnSelectInterface = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.lblInterfaceTag = new System.Windows.Forms.Label();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.button3 = new System.Windows.Forms.Button();
            this.txbDllName = new System.Windows.Forms.TextBox();
            this.cmdOffice = new System.Windows.Forms.ComboBox();
            this.button2 = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOpenFolder
            // 
            this.btnOpenFolder.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOpenFolder.Location = new System.Drawing.Point(768, 315);
            this.btnOpenFolder.Margin = new System.Windows.Forms.Padding(5);
            this.btnOpenFolder.Name = "btnOpenFolder";
            this.btnOpenFolder.Size = new System.Drawing.Size(149, 38);
            this.btnOpenFolder.TabIndex = 20;
            this.btnOpenFolder.Text = "打开所在文件夹";
            this.btnOpenFolder.UseVisualStyleBackColor = true;
            // 
            // txbOfficeFile
            // 
            this.txbOfficeFile.Location = new System.Drawing.Point(5, 151);
            this.txbOfficeFile.Margin = new System.Windows.Forms.Padding(5);
            this.txbOfficeFile.Multiline = true;
            this.txbOfficeFile.Name = "txbOfficeFile";
            this.txbOfficeFile.Size = new System.Drawing.Size(753, 46);
            this.txbOfficeFile.TabIndex = 16;
            // 
            // txbPdf
            // 
            this.txbPdf.Location = new System.Drawing.Point(7, 304);
            this.txbPdf.Margin = new System.Windows.Forms.Padding(5);
            this.txbPdf.Multiline = true;
            this.txbPdf.Name = "txbPdf";
            this.txbPdf.Size = new System.Drawing.Size(751, 61);
            this.txbPdf.TabIndex = 19;
            // 
            // btnSelectOfficeFile
            // 
            this.btnSelectOfficeFile.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSelectOfficeFile.Location = new System.Drawing.Point(768, 151);
            this.btnSelectOfficeFile.Margin = new System.Windows.Forms.Padding(5);
            this.btnSelectOfficeFile.Name = "btnSelectOfficeFile";
            this.btnSelectOfficeFile.Size = new System.Drawing.Size(149, 38);
            this.btnSelectOfficeFile.TabIndex = 17;
            this.btnSelectOfficeFile.Text = "选择 Office 文件";
            this.btnSelectOfficeFile.UseVisualStyleBackColor = true;
            this.btnSelectOfficeFile.Click += new System.EventHandler(this.btnSelectOfficeFile_Click);
            // 
            // btnToPdf
            // 
            this.btnToPdf.Location = new System.Drawing.Point(5, 226);
            this.btnToPdf.Margin = new System.Windows.Forms.Padding(5);
            this.btnToPdf.Name = "btnToPdf";
            this.btnToPdf.Size = new System.Drawing.Size(229, 47);
            this.btnToPdf.TabIndex = 18;
            this.btnToPdf.Text = "转为 pdf";
            this.btnToPdf.UseVisualStyleBackColor = true;
            this.btnToPdf.Click += new System.EventHandler(this.btnToPdf_Click);
            // 
            // txbInterface
            // 
            this.txbInterface.Location = new System.Drawing.Point(5, 31);
            this.txbInterface.Margin = new System.Windows.Forms.Padding(5);
            this.txbInterface.Multiline = true;
            this.txbInterface.Name = "txbInterface";
            this.txbInterface.Size = new System.Drawing.Size(753, 47);
            this.txbInterface.TabIndex = 21;
            // 
            // btnSelectInterface
            // 
            this.btnSelectInterface.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSelectInterface.Location = new System.Drawing.Point(768, 36);
            this.btnSelectInterface.Margin = new System.Windows.Forms.Padding(5);
            this.btnSelectInterface.Name = "btnSelectInterface";
            this.btnSelectInterface.Size = new System.Drawing.Size(149, 38);
            this.btnSelectInterface.TabIndex = 22;
            this.btnSelectInterface.Text = "选择接口程序";
            this.btnSelectInterface.UseVisualStyleBackColor = true;
            this.btnSelectInterface.Click += new System.EventHandler(this.btnSelectInterface_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 83);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 16);
            this.label1.TabIndex = 23;
            this.label1.Text = "接口标识";
            // 
            // lblInterfaceTag
            // 
            this.lblInterfaceTag.AutoSize = true;
            this.lblInterfaceTag.ForeColor = System.Drawing.Color.Blue;
            this.lblInterfaceTag.Location = new System.Drawing.Point(88, 83);
            this.lblInterfaceTag.Name = "lblInterfaceTag";
            this.lblInterfaceTag.Size = new System.Drawing.Size(72, 16);
            this.lblInterfaceTag.TabIndex = 24;
            this.lblInterfaceTag.Text = "接口标识";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.ItemSize = new System.Drawing.Size(76, 38);
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(960, 440);
            this.tabControl1.TabIndex = 25;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.txbInterface);
            this.tabPage1.Controls.Add(this.txbPdf);
            this.tabPage1.Controls.Add(this.btnOpenFolder);
            this.tabPage1.Controls.Add(this.lblInterfaceTag);
            this.tabPage1.Controls.Add(this.txbOfficeFile);
            this.tabPage1.Controls.Add(this.btnToPdf);
            this.tabPage1.Controls.Add(this.btnSelectOfficeFile);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.btnSelectInterface);
            this.tabPage1.Location = new System.Drawing.Point(4, 42);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(952, 394);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "接口转 pdf";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.textBox1);
            this.tabPage2.Controls.Add(this.button1);
            this.tabPage2.Location = new System.Drawing.Point(4, 42);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(952, 394);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "其他";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(43, 83);
            this.textBox1.Margin = new System.Windows.Forms.Padding(5);
            this.textBox1.Multiline = true;
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(779, 209);
            this.textBox1.TabIndex = 24;
            // 
            // button1
            // 
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Location = new System.Drawing.Point(43, 18);
            this.button1.Margin = new System.Windows.Forms.Padding(5);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(181, 38);
            this.button1.TabIndex = 23;
            this.button1.Text = "检测安装的 office";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.button3);
            this.tabPage3.Controls.Add(this.txbDllName);
            this.tabPage3.Controls.Add(this.cmdOffice);
            this.tabPage3.Controls.Add(this.button2);
            this.tabPage3.Location = new System.Drawing.Point(4, 42);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(952, 394);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "接口转 pdf 自动检测";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(649, 51);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(117, 42);
            this.button3.TabIndex = 32;
            this.button3.Text = "指定执行";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // txbDllName
            // 
            this.txbDllName.Location = new System.Drawing.Point(416, 61);
            this.txbDllName.Name = "txbDllName";
            this.txbDllName.Size = new System.Drawing.Size(227, 26);
            this.txbDllName.TabIndex = 31;
            // 
            // cmdOffice
            // 
            this.cmdOffice.FormattingEnabled = true;
            this.cmdOffice.Location = new System.Drawing.Point(258, 61);
            this.cmdOffice.Name = "cmdOffice";
            this.cmdOffice.Size = new System.Drawing.Size(141, 24);
            this.cmdOffice.TabIndex = 30;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(106, 46);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(112, 42);
            this.button2.TabIndex = 29;
            this.button2.Text = "一键执行";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(960, 440);
            this.Controls.Add(this.tabControl1);
            this.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "zsbApp.OfficeToPdf.Test";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnOpenFolder;
        private System.Windows.Forms.TextBox txbOfficeFile;
        private System.Windows.Forms.TextBox txbPdf;
        private System.Windows.Forms.Button btnSelectOfficeFile;
        private System.Windows.Forms.Button btnToPdf;
        private System.Windows.Forms.TextBox txbInterface;
        private System.Windows.Forms.Button btnSelectInterface;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblInterfaceTag;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox txbDllName;
        private System.Windows.Forms.ComboBox cmdOffice;
        private System.Windows.Forms.Button button2;
    }
}

