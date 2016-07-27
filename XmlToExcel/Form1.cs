using BTulz.ModelsTransformer.Transformer;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace XmlToExcel
{
    public partial class Form1 : Form
    {
        List<FolderInfo> _folders;
        List<FolderInfo> folders
        {
            get
            {
                if (_folders == null)
                    _folders = new List<FolderInfo>();
                return _folders;
            }
            set
            {
                _folders = value;
            }
        }
        public Form1()
        {
            InitializeComponent();
        }

        private void btn_input_Click(object sender, EventArgs e)
        {
            fbDialog.Description = "请选择文件夹路径";
            if (fbDialog.ShowDialog() == DialogResult.OK)
            {
                string foldPath = fbDialog.SelectedPath;
                this.txt_input.Text = foldPath;
            }
            this.txt_input.Enabled = true;
            folders.Clear();
        }

        private void btn_output_Click(object sender, EventArgs e)
        {
            fbDialog.Description = "请选择文件夹路径";
            if (fbDialog.ShowDialog() == DialogResult.OK)
            {
                string foldPath = fbDialog.SelectedPath;
                this.txt_output.Text = foldPath;
            }
        }

        private void btn_Get_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(this.txt_input.Text))
                {
                    MessageBox.Show("请选择导入文件夹路径");
                    return;
                }
                folders.Clear();
                DirectoryInfo di = new DirectoryInfo(this.txt_input.Text);
                var childrenDI = di.GetDirectories();//获取子文件夹列表
                foreach (var item in childrenDI)
                {
                    if (Directory.Exists(Path.Combine(item.FullName, "DataStructures")))
                    {
                        FolderInfo folder = new FolderInfo();
                        folder.Selected = true;
                        folder.FolderName = item.Name;
                        folder.FolderPath = Path.Combine(item.FullName, "DataStructures");
                        folders.Add(folder);
                    }
                }
                this.dataGridView1.DataSource = folders;
                this.txt_input.Enabled = false;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }

        private void btn_Export_Click(object sender, EventArgs e)
        {
            try
            {
                string outFolder = this.txt_output.Text;
                if (string.IsNullOrEmpty(outFolder))
                {
                    MessageBox.Show("请选择导出文件夹路径");
                    return;
                }
                foreach (var item in folders.Where(c => c.Selected))
                {
                    try
                    {
                        this.stateLable.Text = string.Format("{0}正在导出...", item.FolderName);
                        Application.DoEvents();
                        DirectoryInfo di = new DirectoryInfo(item.FolderPath);
                        var files = di.GetFiles("*.xml");
                        if (files != null)
                        {
                            string[] fileNames = files.Select(c => c.FullName).ToArray();
                            XmlTransformer myXml = new XmlTransformer();
                            var domainModel = myXml.ToDomainModel(fileNames);
                            XlsTransformerWithToFile myXls = new XlsTransformerWithToFile();
                            myXls.isUseMicrosoftOffice = this.checkBox1.Checked;
                            myXls.folderPath = Path.Combine(Environment.CurrentDirectory, "ExcelMapping");
                            myXls.ToFile(outFolder, domainModel);
                            this.stateLable.Text = string.Format("{0}导出完成", item.FolderName);
                            Application.DoEvents();
                        }
                        else
                        {
                            this.stateLable.Text = string.Format("{0}不需要导出,跳过", item.FolderName);
                            Application.DoEvents();
                        }
                    }
                    catch (Exception error)
                    {
                        MessageBox.Show(error.Message);
                    }
                }
                MessageBox.Show("导出成功");
            }
            catch (Exception error)
            {
                MessageBox.Show(error.Message);
            }
        }


    }

    public class FolderInfo
    {
        public bool Selected { get; set; }
        public string FolderName { get; set; }
        public string FolderPath { get; set; }


    }
}
