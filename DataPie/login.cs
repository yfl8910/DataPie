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
using System.Net;


namespace DataPie
{
    public partial class login : Form
    {
        public static FormMain main = null;
        private DBConfig _DBConfig = new DBConfig();
        string _conString;

        public login()
        {
            //DateTime dt = new DateTime(2015, 1, 1);
            //if (DateTime.Now > dt)
            //{
            //    MessageBox.Show(" 该版本已经太旧了 \r\n 请联系作者更换新版本 \r\n 邮箱：yfl8910@qq.com ");

            //    Application.Exit();
            //    System.Environment.Exit(0);
            //}
            InitializeComponent();
        }

        private void checkDb()
        {
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
        
        }
        //测试连接
        private void btnOk_Click(object sender, EventArgs e)
        {

            checkDb();
            IList<string> _DataBaseList = new List<string>();
             _DBConfig.ServerName = cboServerName.Text.ToString();
            
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
          
        }

        private void login_Load(object sender, EventArgs e)
        {
            DataPieOnLoad();
        }

        private void DataPieOnLoad()
        {
            GetLocalServerIP();
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
        }

        private void btnBrwse_Click(object sender, EventArgs e)
        {

            OpenFileDialog opeanfile = new OpenFileDialog();
            opeanfile.Filter = ("access或者SQLite|*.accdb;*.db");

            opeanfile.RestoreDirectory = true;
            opeanfile.FilterIndex = 1;
            if (opeanfile.ShowDialog() == DialogResult.OK)
            {            
                this.txtconn.Text = opeanfile.FileName;
                txtconn.ReadOnly = true;
            }
        }

        private void btnACC_Click(object sender, EventArgs e)
        {

            string fileName = txtconn.Text.ToString();
            if (txtconn.Text.ToString() == "" || !(System.IO.Path.GetExtension(txtconn.Text.ToString()).ToLower() == ".accdb" || System.IO.Path.GetExtension(txtconn.Text.ToString()).ToLower() == ".db"))
            {
                MessageBox.Show("请选择ACCESS2007或者SQLite数据库！");

                return;
            }
            string ext = fileName.Substring(fileName.LastIndexOf(".") + 1, fileName.Length - fileName.LastIndexOf(".") - 1);

            _DBConfig.DataBase = this.txtconn.Text.ToString();

            if (ext == "accdb")
            {
                //IList<string> _DataBaseList = new List<string>();
                _DBConfig.ProviderName = "ACC";
                _conString = GetConstring(_DBConfig);
                _DBConfig.DBProvider = new DBUtility.DbHelperOleDb(_conString);
             
            }
            else
            {
                //IList<string> _DataBaseList = new List<string>();
                _DBConfig.ProviderName = "SQLite";
                _conString = GetConstring(_DBConfig);
                _DBConfig.DBProvider = new DBUtility.DbHelperSQLite(_conString);
             

            }
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
            else if (sc.Status != System.ServiceProcess.ServiceControllerStatus.Running)
            {
                sc.Start();
                MessageBox.Show("SQL数据库服务启动成功！", "提示信息");
            }



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
            sb.Append(";Initial Catalog=master ;");
            if (db.ValidataType == "Windows身份认证")
            {
                sb.Append(" Integrated Security=SSPI;");
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
                sb.Append(";Initial Catalog=" + db.DataBase + " ; ");
                if (db.ValidataType == "Windows身份认证")
                {
                    sb.Append("Integrated Security=SSPI;Connect Timeout=10000");
                }
                else
                {

                    sb.Append("User ID=" + db.UserName + ";Password=" + db.UserPwd + ";");

                }
                return sb.ToString();
            }
            else if (db.ProviderName == "ACC")
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("Provider=Microsoft.Ace.OleDb.12.0");
                sb.Append(";Data Source= " + db.DataBase);
                sb.Append(";Persist Security Info=False;");
                return sb.ToString();
            }
            else if (db.ProviderName == "SQLite")
            {
                StringBuilder sb = new StringBuilder();
                sb.Append("Data Source= " + db.DataBase + ";");
                return sb.ToString();
            }
            else return "";

        }


        //不再支持ORACEL数据库
        //private void btnOracLogin_Click(object sender, EventArgs e)
        //{
        //    if (textBox1.Text == "" || textBox2.Text == "" || comboBox2.Text == "")
        //    {
        //        MessageBox.Show("用户名、名称、登陆的数据库不能为空！");
        //        return;
        //    }
        //    else
        //    {
        //        string connectionString = "User Id=" + textBox1.Text.ToString().Trim() + ";Password=" + textBox2.Text.ToString().Trim() +
        //      ";Data Source=" + comboBox2.Text.ToString().Trim();
        //        _DBConfig.DBProvider = new DBUtility.DbHelperOra(connectionString);
        //        MainfromShow();
        //        this.Hide();

        //    }
        //}

      
     
        public void GetLocalServerIP()
        {
            cboServerName.Items.Clear();
            cboServerName.Items.Add("127.0.0.1");
            string strHostName = Dns.GetHostName();   //得到本机的主机名
            IPHostEntry ipEntry = Dns.GetHostEntry(strHostName); //取得本机IP

            for (int i = 0; i < ipEntry.AddressList.Count(); i++)
            {
                cboServerName.Items.Add(ipEntry.AddressList[i].ToString());
            }
        }



    }
}