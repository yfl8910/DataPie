using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

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
                int time = UiServices.ExportExcel("Sheet1", sqlText.Text.ToString());
                toolStripStatusLabel1.Text = string.Format("导出的时间为:{0}秒", time);
                toolStripStatusLabel1.ForeColor = Color.Red;
                GC.Collect();
              
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }


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

        //private static string BuildQuery(TableSchema ts)
        //{
        //    StringBuilder sb = new StringBuilder();
        //    sb.Append("SELECT ");
        //    for (int i = 0; i < ts.Columns.Count; i++)
        //    {
        //        sb.Append("[" + ts.Columns[i].ColumnName + "]");
        //        if (i < ts.Columns.Count - 1)
        //            sb.Append(", ");
        //    } 
        //    sb.Append(" FROM " + ts.TableSchemaName + "." + "[" + ts.TableName + "]");
        //    return sb.ToString();
        //}
    }
}
