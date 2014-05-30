namespace DataPie
{
    partial class login
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.btnOk = new System.Windows.Forms.Button();
            this.txtPassword = new System.Windows.Forms.TextBox();
            this.txtUser = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cboValidataType = new System.Windows.Forms.ComboBox();
            this.cboDataBase = new System.Windows.Forms.ComboBox();
            this.cboServerName = new System.Windows.Forms.ComboBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnsel = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnBrwse = new System.Windows.Forms.Button();
            this.btnACC = new System.Windows.Forms.Button();
            this.txtconn = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnOracLogin = new System.Windows.Forms.Button();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.关于ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.退出ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.BackColor = System.Drawing.Color.Transparent;
            this.label6.Location = new System.Drawing.Point(29, 135);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(41, 12);
            this.label6.TabIndex = 93;
            this.label6.Text = "数据库";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(28, 28);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 12);
            this.label5.TabIndex = 88;
            this.label5.Text = "身份验证";
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(92, 214);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(94, 41);
            this.btnOk.TabIndex = 90;
            this.btnOk.Text = "测试连接";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // txtPassword
            // 
            this.txtPassword.Location = new System.Drawing.Point(113, 97);
            this.txtPassword.Name = "txtPassword";
            this.txtPassword.PasswordChar = '*';
            this.txtPassword.Size = new System.Drawing.Size(201, 21);
            this.txtPassword.TabIndex = 85;
            // 
            // txtUser
            // 
            this.txtUser.Location = new System.Drawing.Point(113, 73);
            this.txtUser.Name = "txtUser";
            this.txtUser.Size = new System.Drawing.Size(201, 21);
            this.txtUser.TabIndex = 84;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Location = new System.Drawing.Point(28, 51);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(41, 12);
            this.label4.TabIndex = 82;
            this.label4.Text = "服务器";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(28, 100);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 12);
            this.label3.TabIndex = 81;
            this.label3.Text = "密码";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(28, 76);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 79;
            this.label2.Text = "用户名";
            // 
            // cboValidataType
            // 
            this.cboValidataType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboValidataType.FormattingEnabled = true;
            this.cboValidataType.Items.AddRange(new object[] {
            "Windows身份认证",
            "SQL Server身份认证"});
            this.cboValidataType.Location = new System.Drawing.Point(113, 25);
            this.cboValidataType.Margin = new System.Windows.Forms.Padding(2);
            this.cboValidataType.Name = "cboValidataType";
            this.cboValidataType.Size = new System.Drawing.Size(201, 20);
            this.cboValidataType.TabIndex = 95;
            this.cboValidataType.SelectedIndexChanged += new System.EventHandler(this.cboValidataType_SelectedIndexChanged);
            // 
            // cboDataBase
            // 
            this.cboDataBase.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboDataBase.Enabled = false;
            this.cboDataBase.FormattingEnabled = true;
            this.cboDataBase.Location = new System.Drawing.Point(113, 127);
            this.cboDataBase.Margin = new System.Windows.Forms.Padding(2);
            this.cboDataBase.Name = "cboDataBase";
            this.cboDataBase.Size = new System.Drawing.Size(201, 20);
            this.cboDataBase.TabIndex = 96;
            // 
            // cboServerName
            // 
            this.cboServerName.FormattingEnabled = true;
            this.cboServerName.Items.AddRange(new object[] {
            "Local(本机)"});
            this.cboServerName.Location = new System.Drawing.Point(113, 51);
            this.cboServerName.Margin = new System.Windows.Forms.Padding(2);
            this.cboServerName.Name = "cboServerName";
            this.cboServerName.Size = new System.Drawing.Size(201, 20);
            this.cboServerName.TabIndex = 97;
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(27, 28);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(505, 303);
            this.tabControl1.TabIndex = 99;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.button3);
            this.tabPage1.Controls.Add(this.button2);
            this.tabPage1.Controls.Add(this.button1);
            this.tabPage1.Controls.Add(this.groupBox2);
            this.tabPage1.Controls.Add(this.btnsel);
            this.tabPage1.Controls.Add(this.btnOk);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(497, 277);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "SQL server";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(395, 180);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(93, 28);
            this.button3.TabIndex = 104;
            this.button3.Text = "搜素服务器";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(395, 246);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(96, 25);
            this.button2.TabIndex = 103;
            this.button2.Text = "停止SQL服务";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(395, 214);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(93, 25);
            this.button1.TabIndex = 102;
            this.button1.Text = "开启SQL服务";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cboValidataType);
            this.groupBox2.Controls.Add(this.txtPassword);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.txtUser);
            this.groupBox2.Controls.Add(this.cboServerName);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.cboDataBase);
            this.groupBox2.Location = new System.Drawing.Point(16, 16);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(373, 172);
            this.groupBox2.TabIndex = 101;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "MSSQL 数据库登录";
            // 
            // btnsel
            // 
            this.btnsel.Location = new System.Drawing.Point(224, 214);
            this.btnsel.Name = "btnsel";
            this.btnsel.Size = new System.Drawing.Size(88, 41);
            this.btnsel.TabIndex = 100;
            this.btnsel.Text = "登录数据库";
            this.btnsel.UseVisualStyleBackColor = true;
            this.btnsel.Click += new System.EventHandler(this.btnsel_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.groupBox1);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(497, 277);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "ACCESS或SQLite";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnBrwse);
            this.groupBox1.Controls.Add(this.btnACC);
            this.groupBox1.Controls.Add(this.txtconn);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Location = new System.Drawing.Point(35, 16);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(405, 190);
            this.groupBox1.TabIndex = 102;
            this.groupBox1.TabStop = false;
            // 
            // btnBrwse
            // 
            this.btnBrwse.Location = new System.Drawing.Point(18, 114);
            this.btnBrwse.Name = "btnBrwse";
            this.btnBrwse.Size = new System.Drawing.Size(87, 42);
            this.btnBrwse.TabIndex = 101;
            this.btnBrwse.Text = "浏览...";
            this.btnBrwse.UseVisualStyleBackColor = true;
            this.btnBrwse.Click += new System.EventHandler(this.btnBrwse_Click);
            // 
            // btnACC
            // 
            this.btnACC.Location = new System.Drawing.Point(155, 114);
            this.btnACC.Name = "btnACC";
            this.btnACC.Size = new System.Drawing.Size(92, 42);
            this.btnACC.TabIndex = 100;
            this.btnACC.Text = "登录";
            this.btnACC.UseVisualStyleBackColor = true;
            this.btnACC.Click += new System.EventHandler(this.btnACC_Click);
            // 
            // txtconn
            // 
            this.txtconn.Location = new System.Drawing.Point(123, 20);
            this.txtconn.Multiline = true;
            this.txtconn.Name = "txtconn";
            this.txtconn.Size = new System.Drawing.Size(265, 42);
            this.txtconn.TabIndex = 3;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(16, 33);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(101, 12);
            this.label8.TabIndex = 0;
            this.label8.Text = "数据库文件路径：\r\n";
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.groupBox3);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(497, 277);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "ORACLE";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnOracLogin);
            this.groupBox3.Controls.Add(this.comboBox2);
            this.groupBox3.Controls.Add(this.label9);
            this.groupBox3.Controls.Add(this.textBox1);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Controls.Add(this.textBox2);
            this.groupBox3.Controls.Add(this.label7);
            this.groupBox3.Location = new System.Drawing.Point(43, 24);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(367, 216);
            this.groupBox3.TabIndex = 0;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "ORACLE 数据库登录";
            // 
            // btnOracLogin
            // 
            this.btnOracLogin.Location = new System.Drawing.Point(229, 148);
            this.btnOracLogin.Name = "btnOracLogin";
            this.btnOracLogin.Size = new System.Drawing.Size(95, 41);
            this.btnOracLogin.TabIndex = 93;
            this.btnOracLogin.Tag = "4";
            this.btnOracLogin.Text = "登陆数据库";
            this.btnOracLogin.UseVisualStyleBackColor = true;
            this.btnOracLogin.Click += new System.EventHandler(this.btnOracLogin_Click);
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "ORCL"});
            this.comboBox2.Location = new System.Drawing.Point(123, 99);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(201, 20);
            this.comboBox2.TabIndex = 92;
            this.comboBox2.Tag = "3";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(40, 109);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(41, 12);
            this.label9.TabIndex = 92;
            this.label9.Text = "数据库";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(123, 36);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(201, 21);
            this.textBox1.TabIndex = 90;
            this.textBox1.Tag = "1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(38, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(41, 12);
            this.label1.TabIndex = 86;
            this.label1.Text = "用户名";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(123, 72);
            this.textBox2.Name = "textBox2";
            this.textBox2.PasswordChar = '*';
            this.textBox2.Size = new System.Drawing.Size(201, 21);
            this.textBox2.TabIndex = 91;
            this.textBox2.Tag = "2";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(38, 72);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(29, 12);
            this.label7.TabIndex = 87;
            this.label7.Text = "密码";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.关于ToolStripMenuItem,
            this.退出ToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(570, 25);
            this.menuStrip1.TabIndex = 100;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // 关于ToolStripMenuItem
            // 
            this.关于ToolStripMenuItem.Name = "关于ToolStripMenuItem";
            this.关于ToolStripMenuItem.Size = new System.Drawing.Size(44, 21);
            this.关于ToolStripMenuItem.Text = "关于";
            this.关于ToolStripMenuItem.Click += new System.EventHandler(this.关于ToolStripMenuItem_Click);
            // 
            // 退出ToolStripMenuItem
            // 
            this.退出ToolStripMenuItem.Name = "退出ToolStripMenuItem";
            this.退出ToolStripMenuItem.Size = new System.Drawing.Size(44, 21);
            this.退出ToolStripMenuItem.Text = "退出";
            this.退出ToolStripMenuItem.Click += new System.EventHandler(this.退出ToolStripMenuItem_Click);
            // 
            // login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(570, 356);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "login";
            this.Text = "DataPie登录";
            this.Load += new System.EventHandler(this.login_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.TextBox txtPassword;
        private System.Windows.Forms.TextBox txtUser;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cboValidataType;
        private System.Windows.Forms.ComboBox cboDataBase;
        private System.Windows.Forms.ComboBox cboServerName;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button btnsel;
        private System.Windows.Forms.Button btnBrwse;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtconn;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btnACC;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 关于ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 退出ToolStripMenuItem;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.Button btnOracLogin;
        private System.Windows.Forms.Button button3;
    }
}