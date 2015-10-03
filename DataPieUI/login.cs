using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Net;
using DataPie;


namespace DataPieUI
{
    public partial class login : Form
    {
        public static FormMain main = null;
        string _conString;
        int version=202006;

        public login()
        {    
            //强制更新到最新版，有效期一年
            DateTime dt = new DateTime(version / 100+1, version%100, 1);
            if (DateTime.Now > dt)
            {
                MessageBox.Show(" 该版本已经太旧了 \r\n 请联系作者更换新版本 \r\n 邮箱：yfl8910@qq.com ");

                Application.Exit();
                System.Environment.Exit(0);
            }
            InitializeComponent();
        }

        private int checkDb()
        {
            System.ServiceProcess.ServiceController sc = new System.ServiceProcess.ServiceController();
            sc.ServiceName = "MSSQLSERVER";
            if (sc == null)
            {
                MessageBox.Show("您的机器上没有安装SQL SERVER！", "提示信息");
                return -1;
            }
            else if (sc.Status != System.ServiceProcess.ServiceControllerStatus.Running)
            {
                MessageBox.Show("SQL数据库服务未启动，请点击开启SQL服务！", "提示信息");
                return -2;
            }

            return 0;

        }
        //测试连接
        private void btnOk_Click(object sender, EventArgs e)
        {
            if (checkDb() < 0)
            {
                return;
            }
            else
            {
                IList<string> _DataBaseList = new List<string>();

                DBConfig.db.ServerName = cboServerName.Text.ToString();

                if (cboValidataType.Text == "Windows身份认证")
                {
                    DBConfig.db.ValidataType = "Windows身份认证";
                }
                else
                {
                    txtPassword.Enabled = true;
                    txtUser.Enabled = true;
                    DBConfig.db.ValidataType = "SQL Server身份认证";
                    DBConfig.db.UserName = txtUser.Text.ToString();
                    DBConfig.db.UserPwd = txtPassword.Text.ToString();

                }

                DBConfig.db.ProviderName = "SQL";
                _conString = DBConfig.db.GetSQLmasterConstring();

                DBConfig.db.DBProvider = new DataPie.DBUtility.DbHelperSQL(_conString);
                _DataBaseList = DBConfig.db.DBProvider.GetDataBaseInfo();
                if (_DataBaseList.Count > 0)
                {
                    cboDataBase.DataSource = _DataBaseList;
                    cboDataBase.Enabled = true;
                    cboDataBase.SelectedIndex = 0;
                }

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
                main.Show();


            }
            else
            {
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
            DBConfig.db.ProviderName = "SQL";
            DBConfig.db.DataBase = cboDataBase.Text.ToString();
            _conString = DBConfig.db.GetConstring();
            DBConfig.db.DBProvider = new DataPie.DBUtility.DbHelperSQL(_conString);

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

            DBConfig.db.DataBase = this.txtconn.Text.ToString();

            if (ext == "accdb")
            {
                DBConfig.db.ProviderName = "ACC";
                _conString = DBConfig.db.GetConstring();
                DBConfig.db.DBProvider = new DataPie.DBUtility.DbHelperOleDb(_conString);

            }
            else
            {
                DBConfig.db.ProviderName = "SQLite";
                _conString = DBConfig.db.GetConstring();
                DBConfig.db.DBProvider = new DataPie.DBUtility.DbHelperSQLite(_conString);


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



        public void GetLocalServerIP()
        {
            cboServerName.Items.Clear();
            cboServerName.Items.Add("(local)");
            string strHostName = Dns.GetHostName();   //得到本机的主机名
            IPHostEntry ipEntry = Dns.GetHostEntry(strHostName); //取得本机IP

            for (int i = 0; i < ipEntry.AddressList.Count(); i++)
            {
                cboServerName.Items.Add(ipEntry.AddressList[i].ToString());
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            About about = new About();
            about.ShowDialog();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Application.Exit();
            System.Environment.Exit(0);
        }



    }
}
