using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading.Tasks;

namespace DataPie
{
    public partial class FormSQL : Form
    {
        public static DBConfig _DBConfig;
        public FormSQL()
        {
            InitializeComponent();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {

        }

        public void LoadInfo()
        {
            dbTreeView.Nodes.Clear();
            TreeNode Node = new TreeNode();
            Node.Name = "所有表：";
            Node.Text = "所有表：";
            dbTreeView.Nodes.Add(Node);

            IList<string> list = _DBConfig.DBProvider.GetTableInfo();
            foreach (string s in list)
            {
                TreeNode tn = new TreeNode();
                tn.Name = s;
                tn.Text = s;
                IList<string> col = _DBConfig.DBProvider.GetColumnInfo(s);
                foreach (string c in col)
                {
                    TreeNode colnode = new TreeNode();
                    colnode.Name = c;
                    colnode.Text = c;
                    tn.Nodes.Add(colnode);
                }
                dbTreeView.Nodes["所有表："].Nodes.Add(tn);
            }
            sqlText.Text = " select * from ";

        }

        private void dbTreeView_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void FormSQL_Load(object sender, EventArgs e)
        {
            LoadInfo();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            try
            {
                string filename = UiServices.ShowFileDialog("Sheet1");
                toolStripStatusLabel1.Text = "导数中…";
                toolStripStatusLabel1.ForeColor = Color.Red;
                Task t = TaskExport("Sheet1", sqlText.Text.ToString(), filename);
              
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }


        }

        public async Task TaskExport(string TableName,string sql, string filename)
        {

            await Task.Run(() =>
            {
                int time = UiServices.ExportExcel("Sheet1", sql, filename);
                string s = string.Format("单个OpenXML方式导出的时间为:{0}秒", time);
                this.BeginInvoke(new System.EventHandler(ShowMessage), s);
                MessageBox.Show("导数已完成！");
                GC.Collect();
            });

        }

        private void ShowMessage(object o, System.EventArgs e)
        {
            toolStripStatusLabel1.Text = o.ToString();
            toolStripStatusLabel1.ForeColor = Color.Red;
        }


        private void toolStripButton2_Click_1(object sender, EventArgs e)
        {
            try
            {
                BindingSource bs = new BindingSource();
                DataTable dt = _DBConfig.DBProvider.ReturnDataTable(sqlText.Text.ToString()); 
                bs.DataSource = dt;
                gridResults1.DataSource =bs;
                this.bindingNavigator1.BindingSource = bs;
                dt = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void gridResults1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }


    }
}
