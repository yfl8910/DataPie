using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;



namespace DataPie
{
    public partial class login : Form
    {
        public static FormMain main = null;
        private DBConfig _DBConfig= new DBConfig();
        string _conString;

        public login()
        {
            InitializeComponent();
        }
        //测试连接
        private void btnOk_Click(object sender, EventArgs e)
        {
            IList<string> _DataBaseList = new List<string>();
            System.ServiceProcess.ServiceController sc = new System.ServiceProcess.ServiceController();
            sc.ServiceName = "MSSQLSERVER";
            if (sc == null)
            {
                MessageBox.Show("您的机器上没有安装SQL SERVER！", "提示信息");
                return;
            }
            else if (sc.Status != System.ServiceProcess.ServiceControllerStatus.Running)
            {
                MessageBox.Show("SQL数据库服务未启动，请点击开启SQL服务！", "提示信息");
                return;
            }
            //_DBConfig = new DBConfig();
            if (cboServerName.Text == "Local(本机)")
            {
                _DBConfig.ServerName = "(local)";
            }
            else
            {
                _DBConfig.ServerName = cboServerName.Text.ToString();
            }
            if (cboValidataType.Text == "Windows身份认证")
            {
                _DBConfig.ValidataType = "Windows身份认证";
            }
            else
            {
                txtPassword.Enabled = true;
                txtUser.Enabled = true;
                _DBConfig.ValidataType = "SQL Server身份认证";
                _DBConfig.UserName = txtUser.Text.ToString();
                _DBConfig.UserPwd = txtPassword.Text.ToString();

            }
            _conString = GetSQLmasterConstring(_DBConfig);
            _DBConfig.DBProvider = new DBUtility.DbHelperSQL(_conString);
            _DataBaseList = _DBConfig.DBProvider.GetDataBaseInfo();
            if (_DataBaseList.Count > 0)
            {
                cboDataBase.DataSource = _DataBaseList;
                cboDataBase.Enabled = true;
                cboDataBase.SelectedIndex = 0;
            }
            else
            {
                cboDataBase.Enabled = false;
                cboDataBase.Text = "";
                cboDataBase.DataSource = null;
            }
        }

        private void login_Load(object sender, EventArgs e)
        {
            DataPieOnLoad();
        }

        private void DataPieOnLoad()
        {
            cboServerName.SelectedIndex = 0;
            cboValidataType.SelectedIndex = 0;
            txtPassword.Enabled = false;
            txtUser.Enabled = false;
        }
        private void MainfromShow()
        {
            if (main == null)
            {
                main = new FormMain();

                FormMain.db = _DBConfig;
                UiServices.db = _DBConfig;
                main.Show();


            }
            else
            {
                FormMain.db = _DBConfig;
                UiServices.db = _DBConfig;
                main.DataLoad();
                main.Show();

            }
        }
        //sql数据库选择
        private void btnsel_Click(object sender, EventArgs e)
        {
            if (cboDataBase.Text.ToString() == "")
            {
                MessageBox.Show("请选择数据库名称！");
                return;
            }
            _DBConfig.ProviderName = "SQL";
            _DBConfig.DataBase = cboDataBase.Text.ToString();
            _conString = GetConstring(_DBConfig);
            _DBConfig.DBProvider = new DBUtility.DbHelperSQL(_conString);
            MainfromShow();
            this.Hide();
            //Dispose(); 
        }

        private void btnBrwse_Click(object sender, EventArgs e)
        {

            OpenFileDialog opeanfile = new OpenFileDialog();
            opeanfile.Filter = ("access数据库文件(*.accdb)|*.accdb");

            opeanfile.RestoreDirectory = true;
            opeanfile.FilterIndex = 1;
            if (opeanfile.ShowDialog() == DialogResult.OK)
            {

                this.txtconn.Text = opeanfile.FileName;
                txtconn.ReadOnly = true;
            }
        }

        private void btnTestConn_Click(object sender, EventArgs e)
        {
            if (txtconn.Text.ToString() == "" || System.IO.Path.GetExtension(txtconn.Text.ToString()).ToLower() != ".accdb")
            {
                MessageBox.Show("请选择ACCESS2007数据库文件！");

                return;
            }

            //_DBConfig = new DBConfig();
            IList<string> _DataBaseList = new List<string>();
            _DBConfig.DataBase = this.txtconn.Text.ToString();
            _DBConfig.ProviderName = "ACC";
            _conString = GetConstring(_DBConfig);
            _DBConfig.DBProvider = new DBUtility.DbHelperOleDb(_conString);
            MainfromShow();
            this.Hide();

        }



        private void cboValidataType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cboValidataType.Text.ToString() == "Windows身份认证")
            {
                txtPassword.Enabled = false;
                txtUser.Enabled = false;
            }
            else
            {
                txtPassword.Enabled = true;
                txtUser.Enabled = true;
            }
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
            System.Environment.Exit(0);
        }

        private void 关于ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About about = new About();
            about.ShowDialog();

        }

    
        private void button1_Click(object sender, EventArgs e)
        {
            System.ServiceProcess.ServiceController sc = new System.ServiceProcess.ServiceController();
            sc.ServiceName = "MSSQLSERVER";
            if (sc == null)
            {
                MessageBox.Show("您的机器上没有安装SQL SERVER！", "提示信息");
                return;
            }
            //sc.MachineName = "localhost"; 
            else if (sc.Status != System.ServiceProcess.ServiceControllerStatus.Running)
            {
                sc.Start();
                MessageBox.Show("SQL数据库服务启动成功！", "提示信息");
            }

            

        }

    
     
        /// <summary>
        /// 得到本地服务器
        /// </summary>
        /// <returns></returns>
        public static string[] GetLocalSqlServerNamesWithSqlClientFactory()
        {

            DataTable dataSources = SqlClientFactory.Instance.CreateDataSourceEnumerator().GetDataSources();
            DataColumn column2 = dataSources.Columns["ServerName"];
            DataColumn colume = dataSources.Columns["InstanceName"];
            DataRowCollection rows = dataSources.Rows;
            string[] array = new string[rows.Count];
            for (int i = 0; i < array.Length; i++)
            {
                string str2 = rows[i][column2] as string;
                string str = rows[i][colume] as string;
                if (((str == null) || (str.Length == 0)) || ("MSSQLSERVER" == str))
                {
                    array[i] = str2;
                }
                else
                {
                    array[i] = str2 + @"\" + str;
                }
            }
            Array.Sort<string>(array);

            return array;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.ServiceProcess.ServiceController sc = new System.ServiceProcess.ServiceController();
            sc.ServiceName = "MSSQLSERVER";
            if (!sc.Status.Equals(System.ServiceProcess.ServiceControllerStatus.Stopped))
            {
                sc.Stop();
                MessageBox.Show("SQL数据库服务已经关闭！", "提示信息");
            }
        }

        public static string GetSQLmasterConstring(DBConfig db)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("Data Source=" + db.ServerName);
            sb.Append(";Initial Catalog=master");
            if (db.ValidataType == "Windows身份认证")
            {
                sb.Append("; Integrated Security=SSPI;");
            }
            else
            {

                sb.Append("User ID=" + db.UserName + ";Password=" + db.UserPwd + ";");

            }
            return sb.ToString();
        }

        public static string GetConstring(DBConfig db)
        {
            if (db.ProviderName == "SQL")
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("Data Source=" + db.ServerName);
                sb.Append(";Initial Catalog=" + db.DataBase);
                if (db.ValidataType == "Windows身份认证")
                {
                    sb.Append("; Integrated Security=SSPI;");
                }
                else
                {

                    sb.Append("User ID=" + db.UserName + ";Password=" + db.UserPwd + ";");

                }
                return sb.ToString();
            }
            else
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("Provider=Microsoft.Ace.OleDb.12.0");
                sb.Append(";Data Source= " + db.DataBase);
                sb.Append(";Persist Security Info=False;");
                return sb.ToString();
            }

        }



        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnOracLogin_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "" || comboBox2.Text == "")
            {
                MessageBox.Show("用户名、名称、登陆的数据库不能为空！");
                return;
            }
            else 
            {
                string connectionString = "User Id=" + textBox1.Text.ToString().Trim() + ";Password=" + textBox2.Text.ToString().Trim() +
              ";Data Source=" + comboBox2.Text.ToString().Trim();
                _DBConfig.DBProvider = new DBUtility.DbHelperOra(connectionString);
                //IList<string> list = _DBConfig.DB.GetTableInfo();
                //if (list.Count>0)
                //{ MessageBox.Show("登陆成功！"); }
                //comboBox1.DataSource = db.GetViewInfo();
                //MessageBox.Show(DBUtility.DbHelperOra.GetUser_ID());
                //dataGridView1.DataSource = DBUtility.DbHelperOra.GetSchema("views");
                MainfromShow();
                this.Hide();
            
            }
        }

        private void cSVtoEXCEL工具ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CSVtoEXCEL csv = new CSVtoEXCEL();
            csv.Show();
        }

        /// <summary>
        /// 检索本地服务器
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            cboServerName.Items.Clear();
            //调用得到本地服务器方法
            string[] arr = GetLocalSqlServerNamesWithSqlClientFactory();
            if (arr != null)
            {
                for (int i = 0; i < arr.Length; i++)
                {
                    cboServerName.Items.Add(arr[i]);
                }
                MessageBox.Show("服务器列表已更新！", "提示信息");
            }
        }





    }
}