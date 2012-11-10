
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
        private string serverName;
        /// <summary>
        /// 服务器名
        /// </summary>
        public string ServerName
        {
            get { return serverName; }
            set { serverName = value; }
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
    }
}