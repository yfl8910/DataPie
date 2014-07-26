using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;
using DataPie;
using DataPie.Core;
using System.Linq;

namespace DataPieUI
{
    public partial class FormSQL : Form
    {
        private Point pi;
        string tablename;
        IList<string> col = null;

        public FormSQL()
        {
            InitializeComponent();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            try
            {
                BindingSource bs = new BindingSource();
                DataTable dt = DBConfig.db.DBProvider.ReturnDataTable(sqlText.Text.ToString());
                //bs.DataSource = dt.Select().Take(1000).CopyToDataTable();
                bs.DataSource = dt;
                gridResults1.DataSource = bs;
                this.bindingNavigator1.BindingSource = bs;
                dt = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public void LoadInfo()
        {
            dbTreeView.Nodes.Clear();
            TreeNode Node = new TreeNode();
            Node.Name = "所有表：";
            Node.Text = "所有表：";
            dbTreeView.Nodes.Add(Node);

            IList<string> list = DBConfig.db.DBProvider.GetTableInfo();
            foreach (string s in list)
            {
                TreeNode tn = new TreeNode();
                tn.Name = s;
                tn.Text = s;
                IList<string> col = DBConfig.db.DBProvider.GetColumnInfo(s);
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
        private void dbTreeView_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            pi = new Point(e.X, e.Y);
        }

        private void dbTreeView_DoubleClick(object sender, System.EventArgs e)
        {
            TreeNode node = this.dbTreeView.GetNodeAt(pi);
            if (pi.X < node.Bounds.Left || pi.X > node.Bounds.Right)
            {
                //不触发事件   
                return;
            }
            else
            {
                int i = dbTreeView.SelectedNode.GetNodeCount(false);
                tablename = dbTreeView.SelectedNode.Text.ToString();
                if (i > 0 && tablename != "所有表：")
                {
                    col = DBConfig.db.DBProvider.GetColumnInfo(tablename);
                    string sql = BuildQuery(col, tablename);
                    sqlText.Text = sql;
                    toolStripTextBox1.Text = "";
                    toolStripComboBox1.ComboBox.DataSource = col;

                }

            }
        }

        private static string BuildQuery(IList<string> col, string tablename)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT ");
            sb.Append("\r\n");
            for (int i = 0; i < col.Count; i++)
            {
                sb.Append("[" + col[i] + "]");
                if (i < col.Count - 1)
                    sb.Append(", ");
                sb.Append("\r\n");
            }
            sb.Append("FROM " + "[" + tablename + "]");
            return sb.ToString();
        }

        private static string BuildQuery(IList<string> col, string tablename, int top)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT  TOP " + top);
            sb.Append("\r\n");
            for (int i = 0; i < col.Count; i++)
            {
                sb.Append("[" + col[i] + "]");
                if (i < col.Count - 1)
                    sb.Append(", ");
                sb.Append("\r\n");
            }
            sb.Append("FROM " + "[" + tablename + "]");
            return sb.ToString();

        }

        private void FormSQL_Load(object sender, EventArgs e)
        {
            LoadInfo();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            try
            {

                string filename = Common.ShowFileDialog(tablename + "_" + toolStripTextBox1.Text, ".xlsx");
                ExportExcelAsync(tablename, sqlText.Text.ToString(), filename);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }


        }

        public async void ExportExcelAsync(string TableName, string sql, string filename)
        {
            try
            {
                this.BeginInvoke(new System.EventHandler(ShowMessage), "导数中…");
                IDataReader reader = DBConfig.db.DBProvider.ExecuteReader(sql);
                var t = DataPie.Core.DBToExcel.ExportExcelAsync(filename, sql, TableName);
                await t;
                string s = string.Format("导出成功！耗时:{0}秒", t.Result);
                this.BeginInvoke(new System.EventHandler(ShowMessage), s);
                //MessageBox.Show("导数已完成！");
                GC.Collect();
            }
            catch (Exception ee)
            {
                this.BeginInvoke(new System.EventHandler(ShowErr), ee);
                return;
            }


        }

        private void ShowMessage(object o, System.EventArgs e)
        {
            toolStripStatusLabel1.Text = o.ToString();
            toolStripStatusLabel1.ForeColor = Color.Red;
        }
        private void ShowErr(object o, System.EventArgs e)
        {
            toolStripStatusLabel1.Text = "发生错误！";
            toolStripStatusLabel1.ForeColor = Color.Red;
            Exception ee = o as Exception;
            throw ee;
        }


        private void gridResults1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            try
            {
                string filename = Common.ShowFileDialog(tablename +"_"+ toolStripTextBox1.Text, ".csv");
                if (filename != null)
                {
                    WriteCsvFromsql(sqlText.Text.ToString(), filename);
                }

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        public async void WriteCsvFromsql(string sql, string filename)
        {
            try
            {
                this.BeginInvoke(new System.EventHandler(ShowMessage), "导数中…");
                IDataReader reader = DataPie.DBConfig.db.DBProvider.ExecuteReader(sql);
                var t = DataPie.Core.DBToCsv.ExportCsvAsync(reader, filename);
                await t;
                string s = string.Format("导出成功！耗时:{0}秒", t.Result);
                this.BeginInvoke(new System.EventHandler(ShowMessage), s);
                //MessageBox.Show("导数已完成！");
                GC.Collect();
            }

            catch (Exception ee)
            {
                this.BeginInvoke(new System.EventHandler(ShowErr), ee);
                return;
            }

        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            if (toolStripComboBox1.Text != "")
            {
                string where = " where [" + toolStripComboBox1.Text + "] = '" + toolStripTextBox1.Text + "'";
                sqlText.Text = BuildQuery(col, tablename) + where;
            }

        }



        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            if (toolStripComboBox1.Text != "")
            {
                //string where = " where [" + toolStripComboBox1.Text + "] = '" + toolStripTextBox1.Text + "'";
                //sqlText.Text = BuildQuery(col, tablename, 1000) + where;
                sqlText.Text = BuildQuery(col, tablename, 1000);
            }


        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            try
            {
                string filename = Common.ShowFileDialog(tablename + "_" + toolStripTextBox1.Text, ".csv");
                string sql = sqlText.Text;
                IDataReader reader = DBConfig.db.DBProvider.ExecuteReader(sql);
                if (filename != null)
                {
                    ExportCSVAsync(reader, filename, DataPie.DataPieConfig.RowOutCount);
                }

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public async void ExportCSVAsync(IDataReader reader, string filename, int pagesize)
        {

            try
            {
                this.BeginInvoke(new System.EventHandler(ShowMessage), "导数中…");
                var t = DataPie.Core.DBToCsv.ExportCsvAsync(reader, filename, pagesize);
                await t;
                string s = string.Format("导出成功！耗时:{0}秒", t.Result);
                this.BeginInvoke(new System.EventHandler(ShowMessage), s);
                //MessageBox.Show("导数已完成！");
                GC.Collect();

            }
            catch (Exception ee)
            {
                this.BeginInvoke(new System.EventHandler(ShowErr), ee);
                return;
            }



        }

        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }


    }
}
