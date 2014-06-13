using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataPie
{
    /// <summary>
    /// 数据库配置类
    /// </summary>
    public class DBConfig
    {

        public static DBConfig db = new DBConfig();


        public DBConfig()
        {

        }

        private string providerName;
        /// <summary>
        ///数据提供程序
        /// </summary>
        public string ProviderName
        {
            get { return providerName; }
            set { providerName = value; }
        }

        private string serverName;
        /// <summary>
        /// 服务器名
        /// </summary>
        public string ServerName
        {
            get { return serverName; }
            set { serverName = value; }
        }

        private string validataType;
        /// <summary>
        /// 验证类型
        /// </summary>
        public string ValidataType
        {
            get { return validataType; }
            set { validataType = value; }
        }
        private string userName;
        /// <summary>
        /// 用户名
        /// </summary>
        public string UserName
        {
            get { return userName; }
            set { userName = value; }
        }
        private string userPwd;
        /// <summary>
        /// 密码
        /// </summary>
        public string UserPwd
        {
            get { return userPwd; }
            set { userPwd = value; }
        }
        private string dataBase;
        /// <summary>
        /// 数据库
        /// </summary>
        public string DataBase
        {
            get { return dataBase; }
            set { dataBase = value; }
        }

        private string conString;
        /// <summary>
        /// 连接字符串
        /// </summary>
        public string ConString
        {
            get { return conString; }
            set { conString = value; }
        }
        private DBUtility.IDBUtility dbProvider;
        /// <summary>
        /// 数据访问对象
        /// </summary>
        public DBUtility.IDBUtility DBProvider
        {
            get { return dbProvider; }
            set { dbProvider = value; }
        }

        public string GetSQLmasterConstring()
        {
            if (db.ProviderName == "SQL")
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
            return "";
        }

        public string GetConstring()
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



    }



}
