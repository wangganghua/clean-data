namespace AVC_ClareData
{
    partial class DataClean
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DataClean));
            this.butOpen = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.listFileMessage = new System.Windows.Forms.ListView();
            this.columnFileName = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnData = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnDate = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.butDownload = new System.Windows.Forms.Button();
            this.butOpentxtFile = new System.Windows.Forms.Button();
            this.butSelectData = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // butOpen
            // 
            this.butOpen.Location = new System.Drawing.Point(143, 399);
            this.butOpen.Name = "butOpen";
            this.butOpen.Size = new System.Drawing.Size(75, 23);
            this.butOpen.TabIndex = 0;
            this.butOpen.Text = "打开Excel";
            this.butOpen.UseVisualStyleBackColor = true;
            this.butOpen.Click += new System.EventHandler(this.butOpen_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.listFileMessage);
            this.groupBox1.Location = new System.Drawing.Point(3, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(746, 273);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "文件";
            // 
            // listFileMessage
            // 
            this.listFileMessage.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnFileName,
            this.columnData,
            this.columnDate});
            this.listFileMessage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listFileMessage.Location = new System.Drawing.Point(3, 17);
            this.listFileMessage.MultiSelect = false;
            this.listFileMessage.Name = "listFileMessage";
            this.listFileMessage.Size = new System.Drawing.Size(740, 253);
            this.listFileMessage.TabIndex = 0;
            this.listFileMessage.UseCompatibleStateImageBehavior = false;
            this.listFileMessage.View = System.Windows.Forms.View.Details;
            // 
            // columnFileName
            // 
            this.columnFileName.Text = "文件名";
            this.columnFileName.Width = 200;
            // 
            // columnData
            // 
            this.columnData.Text = "文件大小";
            this.columnData.Width = 80;
            // 
            // columnDate
            // 
            this.columnDate.Text = "日期";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.progressBar1);
            this.groupBox2.Location = new System.Drawing.Point(6, 288);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(743, 72);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "进度";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 17);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "文件个数：";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "当前进度：";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(4, 45);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(731, 14);
            this.progressBar1.TabIndex = 0;
            // 
            // butDownload
            // 
            this.butDownload.Location = new System.Drawing.Point(634, 399);
            this.butDownload.Name = "butDownload";
            this.butDownload.Size = new System.Drawing.Size(75, 23);
            this.butDownload.TabIndex = 3;
            this.butDownload.Text = "下载源文件";
            this.butDownload.UseVisualStyleBackColor = true;
            this.butDownload.Visible = false;
            this.butDownload.Click += new System.EventHandler(this.butDownload_Click);
            // 
            // butOpentxtFile
            // 
            this.butOpentxtFile.Location = new System.Drawing.Point(472, 399);
            this.butOpentxtFile.Name = "butOpentxtFile";
            this.butOpentxtFile.Size = new System.Drawing.Size(75, 23);
            this.butOpentxtFile.TabIndex = 4;
            this.butOpentxtFile.Text = "打开Txt";
            this.butOpentxtFile.UseVisualStyleBackColor = true;
            this.butOpentxtFile.Visible = false;
            this.butOpentxtFile.Click += new System.EventHandler(this.butOpentxtFile_Click);
            // 
            // butSelectData
            // 
            this.butSelectData.Location = new System.Drawing.Point(553, 399);
            this.butSelectData.Name = "butSelectData";
            this.butSelectData.Size = new System.Drawing.Size(75, 23);
            this.butSelectData.TabIndex = 5;
            this.butSelectData.Text = "导出数据";
            this.butSelectData.UseVisualStyleBackColor = true;
            this.butSelectData.Visible = false;
            this.butSelectData.Click += new System.EventHandler(this.butSelectData_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(277, 399);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 6;
            this.button1.Text = "关闭";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(121, 443);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(334, 23);
            this.button2.TabIndex = 7;
            this.button2.Text = "打开:保存excel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // DataClean
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(761, 478);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.butSelectData);
            this.Controls.Add(this.butOpentxtFile);
            this.Controls.Add(this.butDownload);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.butOpen);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "DataClean";
            this.Text = "数据清洗工具";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button butOpen;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ListView listFileMessage;
        private System.Windows.Forms.ColumnHeader columnFileName;
        private System.Windows.Forms.ColumnHeader columnData;
        private System.Windows.Forms.ColumnHeader columnDate;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button butDownload;
        private System.Windows.Forms.Button butOpentxtFile;
        private System.Windows.Forms.Button butSelectData;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}

