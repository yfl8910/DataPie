using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DataPie
{
    public partial class CSVtoEXCEL : Form
    {
        public CSVtoEXCEL()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            if (folder.ShowDialog(this) == DialogResult.OK)
            {
                this.textBox2.Text = folder.SelectedPath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox2.Text.ToString() == "" )
            {
                MessageBox.Show("请选择需要导入的文件夹！");
            }
            DataTable[] ds = UiServices.GetDataTableFromCSV(this.textBox2.Text.ToString(),false);
            string filename = UiServices.ShowFileDialog("Book1");
            int time = UiServices.ExportExcel(ds, "Book1", filename);
            GC.Collect();
            toolStripStatusLabel1.Text = string.Format("执行操作时间为:{0}秒", time);
            toolStripStatusLabel1.ForeColor = Color.Red;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox2.Text.ToString() == "")
            {
                MessageBox.Show("请选择需要导入的文件夹！");
            }
            DataTable[] ds = UiServices.GetDataTableFromCSV(this.textBox2.Text.ToString(), true);
            string filename = UiServices.ShowFileDialog("Book1");
            int time = UiServices.ExportExcel(ds, filename);
            GC.Collect();
            toolStripStatusLabel1.Text = string.Format("执行操作时间为:{0}秒", time);
            toolStripStatusLabel1.ForeColor = Color.Red;
        }
    }
}
