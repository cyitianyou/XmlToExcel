﻿namespace XmlToExcel
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.btn_Export = new System.Windows.Forms.Button();
            this.btn_Get = new System.Windows.Forms.Button();
            this.btn_output = new System.Windows.Forms.Button();
            this.txt_output = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_input = new System.Windows.Forms.Button();
            this.txt_input = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column1 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.Column2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.stateLable = new System.Windows.Forms.ToolStripStatusLabel();
            this.fbDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.splitContainer1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox1.Location = new System.Drawing.Point(0, 0);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(1104, 528);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "XmlToExcel";
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer1.Location = new System.Drawing.Point(4, 24);
            this.splitContainer1.Margin = new System.Windows.Forms.Padding(4);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.checkBox1);
            this.splitContainer1.Panel1.Controls.Add(this.btn_Export);
            this.splitContainer1.Panel1.Controls.Add(this.btn_Get);
            this.splitContainer1.Panel1.Controls.Add(this.btn_output);
            this.splitContainer1.Panel1.Controls.Add(this.txt_output);
            this.splitContainer1.Panel1.Controls.Add(this.label2);
            this.splitContainer1.Panel1.Controls.Add(this.btn_input);
            this.splitContainer1.Panel1.Controls.Add(this.txt_input);
            this.splitContainer1.Panel1.Controls.Add(this.label1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.dataGridView1);
            this.splitContainer1.Panel2.Controls.Add(this.statusStrip1);
            this.splitContainer1.Size = new System.Drawing.Size(1096, 500);
            this.splitContainer1.SplitterDistance = 90;
            this.splitContainer1.SplitterWidth = 6;
            this.splitContainer1.TabIndex = 0;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(51, 97);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(152, 21);
            this.checkBox1.TabIndex = 11;
            this.checkBox1.Text = "使用Office组件";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // btn_Export
            // 
            this.btn_Export.Location = new System.Drawing.Point(915, 60);
            this.btn_Export.Margin = new System.Windows.Forms.Padding(4);
            this.btn_Export.Name = "btn_Export";
            this.btn_Export.Size = new System.Drawing.Size(112, 33);
            this.btn_Export.TabIndex = 10;
            this.btn_Export.Text = "导出";
            this.btn_Export.UseVisualStyleBackColor = true;
            this.btn_Export.Click += new System.EventHandler(this.btn_Export_Click);
            // 
            // btn_Get
            // 
            this.btn_Get.Location = new System.Drawing.Point(915, 14);
            this.btn_Get.Margin = new System.Windows.Forms.Padding(4);
            this.btn_Get.Name = "btn_Get";
            this.btn_Get.Size = new System.Drawing.Size(112, 33);
            this.btn_Get.TabIndex = 9;
            this.btn_Get.Text = "Get";
            this.btn_Get.UseVisualStyleBackColor = true;
            this.btn_Get.Click += new System.EventHandler(this.btn_Get_Click);
            // 
            // btn_output
            // 
            this.btn_output.Location = new System.Drawing.Point(842, 60);
            this.btn_output.Margin = new System.Windows.Forms.Padding(4);
            this.btn_output.Name = "btn_output";
            this.btn_output.Size = new System.Drawing.Size(50, 33);
            this.btn_output.TabIndex = 8;
            this.btn_output.Text = "...";
            this.btn_output.UseVisualStyleBackColor = true;
            this.btn_output.Click += new System.EventHandler(this.btn_output_Click);
            // 
            // txt_output
            // 
            this.txt_output.Location = new System.Drawing.Point(164, 63);
            this.txt_output.Margin = new System.Windows.Forms.Padding(4);
            this.txt_output.Name = "txt_output";
            this.txt_output.Size = new System.Drawing.Size(667, 27);
            this.txt_output.TabIndex = 7;
            this.txt_output.Text = "C:\\Users\\czc\\Desktop\\aaa";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(48, 67);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(110, 17);
            this.label2.TabIndex = 6;
            this.label2.Text = "导出文件夹：";
            // 
            // btn_input
            // 
            this.btn_input.Location = new System.Drawing.Point(842, 14);
            this.btn_input.Margin = new System.Windows.Forms.Padding(4);
            this.btn_input.Name = "btn_input";
            this.btn_input.Size = new System.Drawing.Size(50, 33);
            this.btn_input.TabIndex = 5;
            this.btn_input.Text = "...";
            this.btn_input.UseVisualStyleBackColor = true;
            this.btn_input.Click += new System.EventHandler(this.btn_input_Click);
            // 
            // txt_input
            // 
            this.txt_input.Location = new System.Drawing.Point(164, 17);
            this.txt_input.Margin = new System.Windows.Forms.Padding(4);
            this.txt_input.Name = "txt_input";
            this.txt_input.Size = new System.Drawing.Size(667, 27);
            this.txt_input.TabIndex = 4;
            this.txt_input.Text = "F:\\Project\\XmlToExcel\\XmlToExcel\\数据结构";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(48, 21);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(110, 17);
            this.label1.TabIndex = 3;
            this.label1.Text = "导入文件夹：";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column1,
            this.Column2,
            this.Column3});
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(4);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(1096, 382);
            this.dataGridView1.TabIndex = 0;
            // 
            // Column1
            // 
            this.Column1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.Column1.DataPropertyName = "Selected";
            this.Column1.HeaderText = "是否选择";
            this.Column1.Name = "Column1";
            this.Column1.Width = 82;
            // 
            // Column2
            // 
            this.Column2.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells;
            this.Column2.DataPropertyName = "FolderName";
            this.Column2.HeaderText = "子文件夹";
            this.Column2.Name = "Column2";
            this.Column2.ReadOnly = true;
            this.Column2.Width = 87;
            // 
            // Column3
            // 
            this.Column3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Column3.DataPropertyName = "FolderPath";
            this.Column3.HeaderText = "路径";
            this.Column3.Name = "Column3";
            this.Column3.ReadOnly = true;
            // 
            // statusStrip1
            // 
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(22, 22);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.stateLable});
            this.statusStrip1.Location = new System.Drawing.Point(0, 382);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Padding = new System.Windows.Forms.Padding(2, 0, 21, 0);
            this.statusStrip1.Size = new System.Drawing.Size(1096, 22);
            this.statusStrip1.TabIndex = 1;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // stateLable
            // 
            this.stateLable.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.stateLable.Name = "stateLable";
            this.stateLable.Size = new System.Drawing.Size(0, 17);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1104, 528);
            this.Controls.Add(this.groupBox1);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form1";
            this.Text = "Form1";
            this.groupBox1.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.Button btn_input;
        private System.Windows.Forms.TextBox txt_input;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_output;
        private System.Windows.Forms.TextBox txt_output;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_Export;
        private System.Windows.Forms.Button btn_Get;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.FolderBrowserDialog fbDialog;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel stateLable;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column3;
        private System.Windows.Forms.CheckBox checkBox1;


    }
}

