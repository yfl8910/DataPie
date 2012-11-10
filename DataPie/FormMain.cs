using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.IO;
using Kent.Boogaart.KBCsv;
using Excel = Microsoft.Office.Interop.Excel;


namespace DataPie
{
    public partial class FormMain : Form
    {
        public static DBConfig db;
        public static string conString;

        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            DataLoad();
        }
        /// <summary>
        /// 初始化需要导出的表、视图以及运算的存储过程
        /// </summary>
        public void DataLoad()
        {
            treeView1.Nodes.Clear();
            treeView2.Nodes.Clear();
            TreeNode Node = new TreeNode();

            Node.Name = "所有表：";
            Node.Text = "所有表：";
            treeView1.Nodes.Add(Node);

            Node = new TreeNode();
            Node.Name = "所有视图：";
            Node.Text = "所有视图：";
            treeView1.Nodes.Add(Node);
            IList<string> tableList = new List<string>();

            tableList = db.DBProvider.GetTableInfo();
            foreach (string s in tableList)
            {
                TreeNode tn = new TreeNode();
                tn.Name = s;
                tn.Text = s;
                treeView1.Nodes["所有表："].Nodes.Add(tn);
            }
            IList<string> viewList = new List<string>();
            viewList = db.DBProvider.GetViewInfo();
            foreach (string s in viewList)
            {
                TreeNode tn = new TreeNode();
                tn.Name = s;
                tn.Text = s;
                treeView1.Nodes["所有视图："].Nodes.Add(tn);
            }

            Node = new TreeNode();
            Node.Name = "存储过程";
            Node.Text = "存储过程";
            treeView2.Nodes.Add(Node);
              IList<string>  list = db.DBProvider.GetProcInfo();
            foreach (string s in list)
            {
                TreeNode tn = new TreeNode();
                tn.Name = s;
                tn.Text = s;
                treeView2.Nodes["存储过程"].Nodes.Add(tn);
            }

            treeView1.ExpandAll();
            treeView2.ExpandAll();

            IEnumerable<string> totallist = tableList.Union(viewList);
           
          
            //tableList = _DBConfig.DB.GetTableInfo();

            comboBox1.DataSource = tableList;

            comboBox4.DataSource = totallist.ToList();

            listBox1.Items.Clear();
            listBox2.Items.Clear();
            textBox1.Text = "";
            toolStripStatusLabel2.Text = db.DataBase;

        }

        /// <summary>
        /// 文件浏览
        /// </summary>
        private void btnBrwse_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "EXCEL2007文件|*.xlsx|EXCEL2003文件|*.xls|ACCESS2007文件|*.accdb";

            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog.FileName;
            }
        }
        /// <summary>
        /// 导入EXCEL文件
        /// </summary>
        private void btnImport_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.ToString() == "" || comboBox1.Text.ToString() == "")
            {
                MessageBox.Show("请选择需要导入的文件和导入的表名！");
            }

            else
            {
                string tname = comboBox1.Text.ToString();
                IList<string> List = db.DBProvider.GetColumnInfo(tname);

                string fName = textBox1.Text.ToString();
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();

                DataTable dt = UiServices.GetExcelDataTable(fName, comboBox1.Text.ToString());
                try
                {
                    db.DBProvider.SqlBulkCopyImport(List, comboBox1.Text.ToString(), dt);
                    MessageBox.Show("导入成功");
                }
                catch (Exception ee) { throw ee; }


                watch.Stop();
                GC.Collect();
                toolStripStatusLabel1.Text = string.Format("导入的时间为:{0}秒", watch.ElapsedMilliseconds / 1000);
                toolStripStatusLabel1.ForeColor = Color.Red;
            }
        }

        //导出EXCEL模板文件
        private void btnTP_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.ToString() == "")
            {
                MessageBox.Show("请选择需要导出模板的表名！");
            }
            else
            {
               
                string TableName = comboBox1.Text.ToString();
                int time = UiServices.ExportTemplate(TableName);
                toolStripStatusLabel1.Text = string.Format("导出的时间为:{0}秒", time);
                toolStripStatusLabel1.ForeColor = Color.Red;
            }

        }

        //删除数据库中的数据
        private void btnDel_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.ToString() == "")
            {
                MessageBox.Show("请选择需要删除的表名！");
            }
            else
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                string tname = comboBox1.Text.ToString();
                int num = db.DBProvider.ExecuteSql("delete  from  " + tname);
                watch.Stop();
                if (num > 0)
                {
                    MessageBox.Show("删除成功");
                }
                else
                {
                    MessageBox.Show("删除失败");
                }
                toolStripStatusLabel1.Text = string.Format("删除数据所用时间为:{0}秒", watch.ElapsedMilliseconds / 1000);
                toolStripStatusLabel1.ForeColor = Color.Red;

            }
        }

        //导出数据
        private void btnDtout_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count < 1)
            {
                MessageBox.Show("请选择需要导入的表名！");
            }
            else
            {
                IList<string> SheetNames = new List<string>();
                foreach (var item in listBox1.Items)
                {
                    SheetNames.Add(item.ToString());
                }
                int time = UiServices.ExportExcel(SheetNames);
                toolStripStatusLabel1.Text = string.Format("导出的时间为:{0}秒", time);
                toolStripStatusLabel1.ForeColor = Color.Red;
                GC.Collect();
               

            }
        }



        //增加导出表名
        private void btnAddOne_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Contains(treeView1.SelectedNode.Text.ToString()))
            {
                return;
            }
            else if (listBox1.Items.Count > 9)
            {
                MessageBox.Show("最多可以选择10个表格");
            }
            else
            {
                listBox1.Items.Add(treeView1.SelectedNode.Text.ToString());

            }
        }

        //减少导出表名
        private void btnDeleteOne_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex < 0)
            { MessageBox.Show("请选择删除的表"); }
            else
            { 
                listBox1.Items.RemoveAt(listBox1.SelectedIndex);
            }

        }

        private void 登陆ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            login log = new login();
            log.Show();
        }

        private void btnProcExe_Click(object sender, EventArgs e)
        {
            if (listBox2.Items.Count < 1)
            {
                MessageBox.Show("请选择需要运算的存储过程！");
            }
            else
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                DataTable dt = new DataTable();
                toolStripStatusLabel1.Text = "";
                foreach (var item in listBox2.Items)
                {

                    int i = db.DBProvider.RunProcedure(item.ToString());
                    if (i > 0) 
                    {
                        toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + "存储过程:[" + item.ToString() + "]运算成功！" + "\r\n"; 
                    }
                    else 
                    {
                        toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + "存储过程:[" + item.ToString() + "]运算失败！" + "\r\n";
                    }

                }
                watch.Stop();
                toolStripStatusLabel1.Text = toolStripStatusLabel1.Text + string.Format("请求运算时间为:{0}秒", watch.ElapsedMilliseconds / 1000);
                toolStripStatusLabel1.ForeColor = Color.Red;
                GC.Collect();
                MessageBox.Show("请求运算结束");

            }
        }

        private void btnProcAdd_Click(object sender, EventArgs e)
        {
            if (listBox2.Items.Contains(treeView2.SelectedNode.Text.ToString()))
            {
                MessageBox.Show("已选择，请选择其他表格");
            }
            else if (listBox2.Items.Count > 9)
            {
                MessageBox.Show("最多可以选择10个表格");
            }
            else
            {
                listBox2.Items.Add(treeView2.SelectedNode.Text.ToString());
            }
        }

        private void btnProcDel_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedIndex < 0)
            { MessageBox.Show("请选择删除的存储过程"); }
            else
            { listBox2.Items.RemoveAt(listBox2.SelectedIndex); }
        }



        private void 关于ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About about = new About();
            about.ShowDialog();

        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }
        private Point pi;

        private void treeView1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            pi = new Point(e.X, e.Y);
        }

        private void treeView1_DoubleClick(object sender, System.EventArgs e)
        {
            TreeNode node = this.treeView1.GetNodeAt(pi);
            if (pi.X < node.Bounds.Left || pi.X > node.Bounds.Right)
            {
                //不触发事件   
                return;
            }
            else
            {
                int i = treeView1.SelectedNode.GetNodeCount(false);
                if (!listBox1.Items.Contains(treeView1.SelectedNode.Text.ToString()) && i == 0)

                    listBox1.Items.Add(treeView1.SelectedNode.Text.ToString());
            }
        }
        private void treeView2_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            pi = new Point(e.X, e.Y);
        }

        private void treeView2_DoubleClick(object sender, System.EventArgs e)
        {
            TreeNode node = this.treeView2.GetNodeAt(pi);
            if (pi.X < node.Bounds.Left || pi.X > node.Bounds.Right)
            {
                return;
            }
            else
            {
                int i = treeView2.SelectedNode.GetNodeCount(false);
                if (!listBox2.Items.Contains(treeView2.SelectedNode.Text.ToString()) && i == 0)
                    listBox2.Items.Add(treeView2.SelectedNode.Text.ToString());
            }
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
            System.Environment.Exit(0);
        }

        private void FormMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
            System.Environment.Exit(0);
        }


  
        private void listBox2_DoubleClick(object sender, EventArgs e)
        {

            listBox2.Items.RemoveAt(listBox2.SelectedIndex);
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {

            listBox1.Items.RemoveAt(listBox1.SelectedIndex);
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
            if (textBox2.Text.ToString() == "" || comboBox1.Text.ToString() == "")
            {
                MessageBox.Show("请选择需要导入的文件夹和导入的表名！");
                return;
            }
            Stopwatch watch = Stopwatch.StartNew();
            watch.Start();
            DataTable[] dt = UiServices.GetDataTableFromCSV(this.textBox2.Text.ToString());
            string tname = comboBox1.Text.ToString();
            IList<string> List = db.DBProvider.GetColumnInfo(tname);
   
            for (int i = 0; i < dt.Count(); i++)
            {
                try
                {
                    db.DBProvider.SqlBulkCopyImport(List, comboBox1.Text.ToString(), dt[i]);

                }
                catch (Exception ee)
                { 
                    throw ee;
                }
            }
            
            watch.Stop();
            GC.Collect();
            MessageBox.Show("导入成功");
            toolStripStatusLabel1.Text = string.Format("导入的时间为:{0}秒", watch.ElapsedMilliseconds / 1000);
            toolStripStatusLabel1.ForeColor = Color.Red;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            FormSQL F = new FormSQL();
            FormSQL._DBConfig = db;
            F.Show();
        }



        /// <summary>
        /// 分页导出excel，OpenXML
        /// </summary>
        private void button3_Click(object sender, EventArgs e)
        {
            int pagesize = int.Parse(comboBox3.Text.ToString());
            string TableName = comboBox4.Text.ToString();

            int time = UiServices.ExportExcel(TableName, pagesize);
            toolStripStatusLabel1.Text = string.Format("分页OpenXML导出的时间为:{0}秒", time);
            toolStripStatusLabel1.ForeColor = Color.Red;
            GC.Collect();
        }

        /// <summary>
        /// 分页导出excel，office组件
        /// </summary>
        private void button5_Click(object sender, EventArgs e)
        {
          int pagesize = int.Parse(comboBox3.Text.ToString());
          int time = UiServices.ExportOfficeExcel(comboBox4.Text.ToString(), pagesize);
          toolStripStatusLabel1.Text = string.Format("分页office组件导出的时间为:{0}秒", time);
          toolStripStatusLabel1.ForeColor = Color.Red;
        }

        /// <summary>
        /// 单excel，openXML
        /// </summary>
        private void button6_Click(object sender, EventArgs e)
        {
            string TableName = comboBox4.Text.ToString();
            int time = UiServices.ExportExcel(TableName);
            toolStripStatusLabel1.Text = string.Format("OpenXML导出的时间为:{0}秒", time);
            toolStripStatusLabel1.ForeColor = Color.Red;
            GC.Collect();
        }


        /// <summary>
        /// 单excel，office组件
        /// </summary>
        private void button7_Click(object sender, EventArgs e)
        {
            string TableName = comboBox4.Text.ToString();
            int time = UiServices.ExportOfficeExcel(TableName);
            toolStripStatusLabel1.Text = string.Format("office组件导出的时间为:{0}秒", time);
            toolStripStatusLabel1.ForeColor = Color.Red;
            GC.Collect();
        }

        private void cSVtoEXCEL工具ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CSVtoEXCEL csv = new CSVtoEXCEL();
            csv.Show();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string TableName = comboBox4.Text.ToString();
          
            int time = UiServices.WriteDataTableToCsv(TableName);
            toolStripStatusLabel1.Text = string.Format("csv导出的时间为:{0}秒", time);
            toolStripStatusLabel1.ForeColor = Color.Red;
            GC.Collect();

        }


      
        










    }
}
