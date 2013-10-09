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
using System.Threading.Tasks;


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
            IList<string> list = db.DBProvider.GetProcInfo();
            foreach (string s in list)
            {
                TreeNode tn = new TreeNode();
                tn.Name = s;
                tn.Text = s;
                treeView2.Nodes["存储过程"].Nodes.Add(tn);
            }

            treeView1.ExpandAll();
            treeView2.ExpandAll();

            IEnumerable<string> totallist = tableList.Union(new List<string> { "============" }).Union(viewList);
            comboBox1.DataSource = tableList;
            comboBox2.DataSource = totallist.ToList();
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
                string filename = textBox1.Text.ToString();
                toolStripStatusLabel1.Text = "导数中…";
                toolStripStatusLabel1.ForeColor = Color.Red;
                Task t = TaskImport(List, filename, tname);
            }
        }
        //excel异步方式导入
        public async Task TaskImport(IList<string> List, string filename, string tname)
        {

            await Task.Run(() =>
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                try
                {
                    DataTable dt = UiServices.GetExcelDataTable(filename, tname);
                    db.DBProvider.SqlBulkCopyImport(List, tname, dt);

                }
                catch (Exception ee)
                {
                    this.BeginInvoke(new System.EventHandler(ShowErr), ee);
                    return;
                }
                watch.Stop();
                string s = "导入成功！ 使用时间：" + watch.ElapsedMilliseconds / 1000 + "秒";
                this.BeginInvoke(new System.EventHandler(ShowMessage), s);
                MessageBox.Show("导入成功！");
                GC.Collect();
            });

        }

        private void ShowErr(object o, System.EventArgs e)
        {
            toolStripStatusLabel1.Text = "发生错误！";
            toolStripStatusLabel1.ForeColor = Color.Red;
            Exception ee = o as Exception;
            throw ee;
        }


        //csv文件夹导入
        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox2.Text.ToString() == "" || comboBox1.Text.ToString() == "")
            {
                MessageBox.Show("请选择需要导入的文件夹和导入的表名！");
                return;
            }
            string path = this.textBox2.Text.ToString();
            string tname = comboBox1.Text.ToString();
            IList<string> List = db.DBProvider.GetColumnInfo(tname);
            toolStripStatusLabel1.Text = "导数中…";
            toolStripStatusLabel1.ForeColor = Color.Red;
            Task t = TaskImportCsv(List, path, tname);
        }

        //csv异步方式导入
        //public async Task TaskImportCsv(IList<string> List, string path, string tname)
        //{

        //    await Task.Run(() =>
        //    {
        //        Stopwatch watch = Stopwatch.StartNew();
        //        watch.Start();
        //        DataTable[] dt = UiServices.GetDataTableFromCSV(path,false);
        //        for (int i = 0; i < dt.Count(); i++)
        //        {
        //            try
        //            {
        //                db.DBProvider.SqlBulkCopyImport(List, tname, dt[i]);
        //            }
        //            catch (Exception ee)
        //            {
        //                this.BeginInvoke(new System.EventHandler(ShowErr), ee);
        //                return;
        //            }
        //        }
        //        watch.Stop();
        //        string s = "导入成功！";
        //        this.BeginInvoke(new System.EventHandler(ShowMessage), s);
        //        MessageBox.Show("导入成功");
        //        GC.Collect();
        //    });

        //}
        public async Task TaskImportCsv(IList<string> List, string path, string tname)
        {

            await Task.Run(() =>
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                List<FileInfo> filelist = UiServices.GetFilelist(path, false);
                for (int i = 0; i < filelist.Count(); i++)
                {
                    try
                    {
                        string m = "正在导入第 " + (i + 1) + " 个文件：" + filelist[i].ToString();
                        this.BeginInvoke(new System.EventHandler(ShowMessage), m);
                        DataTable dt = UiServices.GetDataTableFromfile(filelist[i]);
                        db.DBProvider.SqlBulkCopyImport(List, tname, dt);
                    }
                    catch (Exception ee)
                    {
                        this.BeginInvoke(new System.EventHandler(ShowErr), ee);
                        return;
                    }
                }
                watch.Stop();
                string s = "导入成功！ 使用时间：" + watch.ElapsedMilliseconds / 1000 + "秒";
                this.BeginInvoke(new System.EventHandler(ShowMessage), s);
                MessageBox.Show("导入成功");
                GC.Collect();
            });

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
                string filename = UiServices.ShowFileDialog(TableName);
                int time = UiServices.ExportTemplate(TableName, filename);
                toolStripStatusLabel1.Text = string.Format("导出的时间为:{0}秒", time);
                toolStripStatusLabel1.ForeColor = Color.Red;
                MessageBox.Show("导出成功！");
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
                toolStripStatusLabel1.Text = "删除数据中…";
                toolStripStatusLabel1.ForeColor = Color.Red;

                string tname = comboBox1.Text.ToString();
                int num = db.DBProvider.ExecuteSql("TRUNCATE TABLE   " + tname);
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

        public static string ShowFileDialog(string FileName)
        {
            System.Windows.Forms.SaveFileDialog saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog1.Filter = "excel2007|*.xlsx";
            saveFileDialog1.FileName = FileName;
            saveFileDialog1.DefaultExt = ".xlsx";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                return saveFileDialog1.FileName.ToString();
            }
            else
            {
                return null;
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
                string filename = UiServices.ShowFileDialog(SheetNames[0]);
                toolStripStatusLabel1.Text = "导数中…";
                toolStripStatusLabel1.ForeColor = Color.Red;
                if (filename != null)
                { Task t = TaskExport(SheetNames, filename); }


            }
        }



        //异步导出EXCEL
        public async Task TaskExport(IList<string> SheetNames, string filename)
        {

            await Task.Run(() =>
              {

                  try
                  {
                      int time = UiServices.ExportExcel(SheetNames, filename);
                      string s = string.Format("导出的时间为:{0}秒", time);
                      this.BeginInvoke(new System.EventHandler(ShowMessage), s);
                      MessageBox.Show("导数已完成！");
                      GC.Collect();

                  }
                  catch (Exception ee)
                  {
                      this.BeginInvoke(new System.EventHandler(ShowErr), ee);
                      return;
                  }

              });




        }

        private void ShowMessage(object o, System.EventArgs e)
        {
            toolStripStatusLabel1.Text = o.ToString();
            toolStripStatusLabel1.ForeColor = Color.Red;
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

        //请求计算事件
        private void btnProcExe_Click(object sender, EventArgs e)
        {
            if (listBox2.Items.Count < 1)
            {
                MessageBox.Show("请选择需要运算的存储过程！");
            }
            else
            {
                IList<string> list = new List<string>();
                foreach (var item in listBox2.Items)
                {
                    list.Add(item.ToString());
                }
                toolStripStatusLabel1.Text = "存储过程计算中…";
                toolStripStatusLabel1.ForeColor = Color.Red;
                Task t = TaskProcExeute(list);
            }
        }

        //异步方式存储过程调用
        public async Task TaskProcExeute(IList<string> procs)
        {

            await Task.Run(() =>
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                string s = "";
                try
                {

                    foreach (var item in procs)
                    {
                        int i = db.DBProvider.RunProcedure(item.ToString());
                        if (i > 0)
                        { s = "存储过程:[" + item.ToString() + "]运算成功！" + "\r\n"; }
                        else
                        { s = "存储过程:[" + item.ToString() + "]运算失败！" + "\r\n"; }
                    }


                }
                catch (Exception ee)
                {
                    this.BeginInvoke(new System.EventHandler(ShowErr), ee);
                    return;
                }

                watch.Stop();

                s = s + string.Format("请求运算时间为:{0}秒", watch.ElapsedMilliseconds / 1000);
                this.BeginInvoke(new System.EventHandler(ShowMessage), s);
                MessageBox.Show("请求运算结束！");
                return;
            });
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

        //选择csv文件夹
        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folder = new FolderBrowserDialog();
            if (folder.ShowDialog(this) == DialogResult.OK)
            {
                this.textBox2.Text = folder.SelectedPath;
            }
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
            string filename = UiServices.ShowFileDialog(TableName);
            toolStripStatusLabel1.Text = "导数中…";
            toolStripStatusLabel1.ForeColor = Color.Red;
            if (filename != null)
            { Task t = TaskExport(TableName, filename, pagesize); }

        }

        //异步导出分页OpenXMLL
        public async Task TaskExport(string TableName, string filename, int pagesize)
        {

            await Task.Run(() =>
            {

                int time = UiServices.ExportExcel(TableName, pagesize, filename);
                string s = string.Format("分页OpenXML方式导出的时间为:{0}秒", time);
                this.BeginInvoke(new System.EventHandler(ShowMessage), s);
                MessageBox.Show("导数已完成！");
                GC.Collect();
            });
        }

        /// <summary>
        /// 单excel，openXML
        /// </summary>
        private void button6_Click(object sender, EventArgs e)
        {
            string TableName = comboBox4.Text.ToString();
            string filename = UiServices.ShowFileDialog(TableName);
            toolStripStatusLabel1.Text = "导数中…";
            toolStripStatusLabel1.ForeColor = Color.Red;
            if (filename != null)
            { Task t = TaskExport(TableName, filename); }

        }


        //异步单openXML方式导出
        public async Task TaskExport(string TableName, string filename)
        {

            await Task.Run(() =>
              {
                  int time = UiServices.ExportExcel(TableName, filename);
                  string s = string.Format("单个OpenXML方式导出的时间为:{0}秒", time);
                  this.BeginInvoke(new System.EventHandler(ShowMessage), s);
                  MessageBox.Show("导数已完成！");
                  GC.Collect();
              });

        }


        private void cSVtoEXCEL工具ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CSVtoEXCEL csv = new CSVtoEXCEL();
            csv.Show();
        }

        /// <summary>
        /// 单个csv
        /// </summary>
        private void button8_Click(object sender, EventArgs e)
        {
            string TableName = comboBox4.Text.ToString();
            System.Windows.Forms.SaveFileDialog saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog1.Filter = "csv文件|*.csv";
            saveFileDialog1.FileName = TableName;
            saveFileDialog1.DefaultExt = ".csv";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && TableName != "")
            {
                string filename = saveFileDialog1.FileName.ToString();
                toolStripStatusLabel1.Text = "导数中…";
                toolStripStatusLabel1.ForeColor = Color.Red;
                Task t = TaskExportCsv(TableName, filename);
            }
        }

        //异步单个csv导出
        public async Task TaskExportCsv(string TableName, string filename)
        {
            await Task.Run(() =>
            {
                int time = UiServices.WriteDataTableToCsv(TableName, filename);
                string s = string.Format("单个csv方式导出的时间为:{0}秒", time);
                this.BeginInvoke(new System.EventHandler(ShowMessage), s);
                MessageBox.Show("导数已完成！");
                GC.Collect();
            });

        }


        /// <summary>
        /// 多个csv
        /// </summary>
        private void button5_Click(object sender, EventArgs e)
        {
            string TableName = comboBox4.Text.ToString();
            int pagesize = int.Parse(comboBox3.Text.ToString());
            System.Windows.Forms.SaveFileDialog saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog1.Filter = "csv文件|*.csv";
            saveFileDialog1.FileName = TableName;
            saveFileDialog1.DefaultExt = ".csv";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK && TableName != "")
            {
                string filename = saveFileDialog1.FileName.ToString();
                toolStripStatusLabel1.Text = "导数中…";
                toolStripStatusLabel1.ForeColor = Color.Red;
                Task t = TaskExportCsv(TableName, filename, pagesize);
            }
        }

        //异步分页导出csv
        public async Task TaskExportCsv(string TableName, string filename, int pagesize)
        {
            await Task.Run(() =>
            {
                int time = UiServices.WriteDataTableToCsv(TableName, pagesize, filename);
                string s = string.Format("分页csv方式导出的时间为:{0}秒", time);
                this.BeginInvoke(new System.EventHandler(ShowMessage), s);
                MessageBox.Show("导数已完成！");
                GC.Collect();
            });

        }

        private void button7_Click(object sender, EventArgs e)
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


                System.Windows.Forms.SaveFileDialog saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
                saveFileDialog1.Filter = "csv文件|*.csv";
                saveFileDialog1.FileName = "output";
                saveFileDialog1.DefaultExt = ".csv";
                if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string filename = saveFileDialog1.FileName.ToString();
                    toolStripStatusLabel1.Text = "导数中…";
                    toolStripStatusLabel1.ForeColor = Color.Red;
                    Task t = TaskExportMuticsv(SheetNames, filename);
                }


            }
        }
        //异步导出多csv
        public async Task TaskExportMuticsv(IList<string> SheetNames, string filename)
        {

            await Task.Run(() =>
            {

                try
                {
                    int time = UiServices.ExportMuticsv(SheetNames, filename);
                    string s = string.Format("导出的时间为:{0}秒", time);
                    this.BeginInvoke(new System.EventHandler(ShowMessage), s);
                    MessageBox.Show("导数已完成！");
                    GC.Collect();

                }
                catch (Exception ee)
                {
                    this.BeginInvoke(new System.EventHandler(ShowErr), ee);
                    return;
                }

            });




        }

        private void button9_Click(object sender, EventArgs e)
        {
            string tbname = comboBox2.Text.ToString();
            string colname = comboBox5.Text.ToString();
            string csql = " select distinct " + "[" + colname + "]" + " from " + "[" + tbname + "]";
            string sql = " select * from  " + "[" + tbname + "]" + " where  " + "[" + colname + "]" + " = ";

            IList<string> clums = new List<string>();
            DataTable dt = UiServices.GetDataTableFromSQL(csql);
            int num = dt.Rows.Count;
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow _DataRowItem in dt.Rows)
                {
                    clums.Add(_DataRowItem[colname].ToString());
                }
            }
            string filename = UiServices.ShowFileDialog(tbname);
            Stopwatch watch = Stopwatch.StartNew();
            watch.Start();
            foreach (var a in clums)
            {

                string savepath = filename.Remove(filename.LastIndexOf('.')) + '_' + a + ".xlsx";

                DataTable dt1 = UiServices.GetDataTableFromSQL(sql + "'" + a + "'");
                UiServices.SaveExcel(savepath, dt1, "sheet1");

            }
            watch.Stop();
            toolStripStatusLabel1.Text = string.Format("导出的时间为:{0}秒", watch.ElapsedMilliseconds / 1000);
            toolStripStatusLabel1.ForeColor = Color.Red;
            MessageBox.Show("导出成功");
            GC.Collect();
        }



        private void getcolumn()
        {
            IList<string> List = db.DBProvider.GetColumnInfo(comboBox2.Text.ToString());
            comboBox5.DataSource = List;

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            getcolumn();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string tbname = comboBox2.Text.ToString();
            string colname = comboBox5.Text.ToString();
            string csql = " select distinct " + "[" + colname + "]" + " from " + "[" + tbname + "]";
            string sql = " select * from  " + "[" + tbname + "]" + " where  " + "[" + colname + "]" + " = ";

            IList<string> clums = new List<string>();
            DataTable dt = UiServices.GetDataTableFromSQL(csql);
            int num = dt.Rows.Count;
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow _DataRowItem in dt.Rows)
                {
                    clums.Add(_DataRowItem[colname].ToString());
                }
            }

            string filename = "";
            System.Windows.Forms.SaveFileDialog saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog1.Filter = "CSV|*.csv";
            saveFileDialog1.FileName = tbname;
            saveFileDialog1.DefaultExt = ".csv";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename = saveFileDialog1.FileName.ToString();
            }
            else { return; }
            //string filename = UiServices.ShowFileDialog(tbname);
            Stopwatch watch = Stopwatch.StartNew();
            watch.Start();
            foreach (var a in clums)
            {

                string savepath = filename.Remove(filename.LastIndexOf('.')) + '_' + a + ".csv";

                DataTable dt1 = UiServices.GetDataTableFromSQL(sql + "'" + a + "'");
                UiServices.SaveCsv(dt1, savepath);

            }
            watch.Stop();
            toolStripStatusLabel1.Text = string.Format("导出的时间为:{0}秒", watch.ElapsedMilliseconds / 1000);
            toolStripStatusLabel1.ForeColor = Color.Red;
            MessageBox.Show("导出成功");
            GC.Collect();
        }



    }
}
