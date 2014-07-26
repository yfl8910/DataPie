using System;
using System.Collections;
using System.Data;
using System.Data.SQLite;
using System.Collections.Generic;
using System.Text;

namespace DataPie.DBUtility
{
    /// <summary>
    ///  
    /// 数据访问基础类(基于SQLite)
    /// 可以用户可以修改满足自己项目的需要。
    /// </summary>
    public class DbHelperSQLite : IDBUtility
    {
        //数据库连接字符串(web.config来配置)，可以动态更改connectionString支持多数据库.		
        public static string connectionString = "";
        public DbHelperSQLite()
        {
        }
        public DbHelperSQLite(string strConnectionString)
        {
            connectionString = strConnectionString;
        }


        #region 公用方法

        public static int GetMaxID(string FieldName, string TableName)
        {
            string strsql = "select max(" + FieldName + ")+1 from " + TableName;
            object obj = GetSingle(strsql);
            if (obj == null)
            {
                return 1;
            }
            else
            {
                return int.Parse(obj.ToString());
            }
        }
        public static bool Exists(string strSql)
        {
            object obj = GetSingle(strSql);
            int cmdresult;
            if ((Object.Equals(obj, null)) || (Object.Equals(obj, System.DBNull.Value)))
            {
                cmdresult = 0;
            }
            else
            {
                cmdresult = int.Parse(obj.ToString());
            }
            if (cmdresult == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        public static bool Exists(string strSql, params SQLiteParameter[] cmdParms)
        {
            object obj = GetSingle(strSql, cmdParms);
            int cmdresult;
            if ((Object.Equals(obj, null)) || (Object.Equals(obj, System.DBNull.Value)))
            {
                cmdresult = 0;
            }
            else
            {
                cmdresult = int.Parse(obj.ToString());
            }
            if (cmdresult == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        #endregion

        #region  执行简单SQL语句

        /// <summary>
        /// 执行SQL语句，返回影响的记录数
        /// </summary>
        /// <param name="SQLString">SQL语句</param>
        /// <returns>影响的记录数</returns>
        public int ExecuteSql(string SQLString)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                using (SQLiteCommand cmd = new SQLiteCommand(SQLString, connection))
                {
                    try
                    {
                        connection.Open();
                        int rows = cmd.ExecuteNonQuery();
                        return rows;
                    }
                    catch (System.Data.SQLite.SQLiteException E)
                    {
                        connection.Close();
                        throw new Exception(E.Message);
                    }
                }
            }
        }

        /// <summary>
        /// 执行多条SQL语句，实现数据库事务。
        /// </summary>
        /// <param name="SQLStringList">多条SQL语句</param>		
        public static void ExecuteSqlTran(ArrayList SQLStringList)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = conn;
                SQLiteTransaction tx = conn.BeginTransaction();
                cmd.Transaction = tx;
                try
                {
                    for (int n = 0; n < SQLStringList.Count; n++)
                    {
                        string strsql = SQLStringList[n].ToString();
                        if (strsql.Trim().Length > 1)
                        {
                            cmd.CommandText = strsql;
                            cmd.ExecuteNonQuery();
                        }
                    }
                    tx.Commit();
                }
                catch (System.Data.SQLite.SQLiteException E)
                {
                    tx.Rollback();
                    throw new Exception(E.Message);
                }
            }
        }
        /// <summary>
        /// 执行带一个存储过程参数的的SQL语句。
        /// </summary>
        /// <param name="SQLString">SQL语句</param>
        /// <param name="content">参数内容,比如一个字段是格式复杂的文章，有特殊符号，可以通过这个方式添加</param>
        /// <returns>影响的记录数</returns>
        public static int ExecuteSql(string SQLString, string content)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                SQLiteCommand cmd = new SQLiteCommand(SQLString, connection);
                SQLiteParameter myParameter = new SQLiteParameter("@content", DbType.String);
                myParameter.Value = content;
                cmd.Parameters.Add(myParameter);
                try
                {
                    connection.Open();
                    int rows = cmd.ExecuteNonQuery();
                    return rows;
                }
                catch (System.Data.SQLite.SQLiteException E)
                {
                    throw new Exception(E.Message);
                }
                finally
                {
                    cmd.Dispose();
                    connection.Close();
                }
            }
        }
        /// <summary>
        /// 向数据库里插入图像格式的字段(和上面情况类似的另一种实例)
        /// </summary>
        /// <param name="strSQL">SQL语句</param>
        /// <param name="fs">图像字节,数据库的字段类型为image的情况</param>
        /// <returns>影响的记录数</returns>
        public static int ExecuteSqlInsertImg(string strSQL, byte[] fs)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                SQLiteCommand cmd = new SQLiteCommand(strSQL, connection);
                SQLiteParameter myParameter = new SQLiteParameter("@fs", DbType.Binary);
                myParameter.Value = fs;
                cmd.Parameters.Add(myParameter);
                try
                {
                    connection.Open();
                    int rows = cmd.ExecuteNonQuery();
                    return rows;
                }
                catch (System.Data.SQLite.SQLiteException E)
                {
                    throw new Exception(E.Message);
                }
                finally
                {
                    cmd.Dispose();
                    connection.Close();
                }
            }
        }

        /// <summary>
        /// 执行一条计算查询结果语句，返回查询结果（object）。
        /// </summary>
        /// <param name="SQLString">计算查询结果语句</param>
        /// <returns>查询结果（object）</returns>
        public static object GetSingle(string SQLString)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                using (SQLiteCommand cmd = new SQLiteCommand(SQLString, connection))
                {
                    try
                    {
                        connection.Open();
                        object obj = cmd.ExecuteScalar();
                        if ((Object.Equals(obj, null)) || (Object.Equals(obj, System.DBNull.Value)))
                        {
                            return null;
                        }
                        else
                        {
                            return obj;
                        }
                    }
                    catch (System.Data.SQLite.SQLiteException e)
                    {
                        connection.Close();
                        throw new Exception(e.Message);
                    }
                }
            }
        }
        /// <summary>
        /// 执行查询语句，返回SQLiteDataReader
        /// </summary>
        /// <param name="strSQL">查询语句</param>
        /// <returns>SQLiteDataReader</returns>
        public  IDataReader ExecuteReader(string strSQL)
        {
            SQLiteConnection connection = new SQLiteConnection(connectionString);
            SQLiteCommand cmd = new SQLiteCommand(strSQL, connection);
            cmd.CommandTimeout = 10000;
            try
            {
                connection.Open();
                SQLiteDataReader myReader = cmd.ExecuteReader();
                return myReader;
            }
            catch (System.Data.SQLite.SQLiteException e)
            {
                throw new Exception(e.Message);
            }

        }
        /// <summary>
        /// 执行查询语句，返回DataSet
        /// </summary>
        /// <param name="SQLString">查询语句</param>
        /// <returns>DataSet</returns>
        public static DataSet Query(string SQLString)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                DataSet ds = new DataSet();
                try
                {
                    connection.Open();
                    SQLiteDataAdapter command = new SQLiteDataAdapter(SQLString, connection);
                    command.Fill(ds, "ds");
                }
                catch (System.Data.SQLite.SQLiteException ex)
                {
                    throw new Exception(ex.Message);
                }
                return ds;
            }
        }


        #endregion

        #region 执行带参数的SQL语句

        /// <summary>
        /// 执行SQL语句，返回影响的记录数
        /// </summary>
        /// <param name="SQLString">SQL语句</param>
        /// <returns>影响的记录数</returns>
        public static int ExecuteSql(string SQLString, params SQLiteParameter[] cmdParms)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                using (SQLiteCommand cmd = new SQLiteCommand())
                {
                    try
                    {
                        PrepareCommand(cmd, connection, null, SQLString, cmdParms);
                        int rows = cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                        return rows;
                    }
                    catch (System.Data.SQLite.SQLiteException E)
                    {
                        throw new Exception(E.Message);
                    }
                }
            }
        }


        /// <summary>
        /// 执行多条SQL语句，实现数据库事务。
        /// </summary>
        /// <param name="SQLStringList">SQL语句的哈希表（key为sql语句，value是该语句的SQLiteParameter[]）</param>
        public static void ExecuteSqlTran(Hashtable SQLStringList)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                using (SQLiteTransaction trans = conn.BeginTransaction())
                {
                    SQLiteCommand cmd = new SQLiteCommand();
                    try
                    {
                        //循环
                        foreach (DictionaryEntry myDE in SQLStringList)
                        {
                            string cmdText = myDE.Key.ToString();
                            SQLiteParameter[] cmdParms = (SQLiteParameter[])myDE.Value;
                            PrepareCommand(cmd, conn, trans, cmdText, cmdParms);
                            int val = cmd.ExecuteNonQuery();
                            cmd.Parameters.Clear();

                            trans.Commit();
                        }
                    }
                    catch
                    {
                        trans.Rollback();
                        throw;
                    }
                }
            }
        }


        /// <summary>
        /// 执行一条计算查询结果语句，返回查询结果（object）。
        /// </summary>
        /// <param name="SQLString">计算查询结果语句</param>
        /// <returns>查询结果（object）</returns>
        public static object GetSingle(string SQLString, params SQLiteParameter[] cmdParms)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                using (SQLiteCommand cmd = new SQLiteCommand())
                {
                    try
                    {
                        PrepareCommand(cmd, connection, null, SQLString, cmdParms);
                        object obj = cmd.ExecuteScalar();
                        cmd.Parameters.Clear();
                        if ((Object.Equals(obj, null)) || (Object.Equals(obj, System.DBNull.Value)))
                        {
                            return null;
                        }
                        else
                        {
                            return obj;
                        }
                    }
                    catch (System.Data.SQLite.SQLiteException e)
                    {
                        throw new Exception(e.Message);
                    }
                }
            }
        }

        /// <summary>
        /// 执行查询语句，返回SQLiteDataReader
        /// </summary>
        /// <param name="strSQL">查询语句</param>
        /// <returns>SQLiteDataReader</returns>
        public static SQLiteDataReader ExecuteReader(string SQLString, params SQLiteParameter[] cmdParms)
        {
            SQLiteConnection connection = new SQLiteConnection(connectionString);
            SQLiteCommand cmd = new SQLiteCommand();
            try
            {
                PrepareCommand(cmd, connection, null, SQLString, cmdParms);
                SQLiteDataReader myReader = cmd.ExecuteReader();
                cmd.Parameters.Clear();
                return myReader;
            }
            catch (System.Data.SQLite.SQLiteException e)
            {
                throw new Exception(e.Message);
            }

        }

        /// <summary>
        /// 执行查询语句，返回DataSet
        /// </summary>
        /// <param name="SQLString">查询语句</param>
        /// <returns>DataSet</returns>
        public static DataSet Query(string SQLString, params SQLiteParameter[] cmdParms)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                SQLiteCommand cmd = new SQLiteCommand();
                PrepareCommand(cmd, connection, null, SQLString, cmdParms);
                using (SQLiteDataAdapter da = new SQLiteDataAdapter(cmd))
                {
                    DataSet ds = new DataSet();
                    try
                    {
                        da.Fill(ds, "ds");
                        cmd.Parameters.Clear();
                    }
                    catch (System.Data.SQLite.SQLiteException ex)
                    {
                        throw new Exception(ex.Message);
                    }
                    return ds;
                }
            }
        }


        private static void PrepareCommand(SQLiteCommand cmd, SQLiteConnection conn, SQLiteTransaction trans, string cmdText, SQLiteParameter[] cmdParms)
        {
            if (conn.State != ConnectionState.Open)
                conn.Open();
            cmd.Connection = conn;
            cmd.CommandText = cmdText;
            if (trans != null)
                cmd.Transaction = trans;
            cmd.CommandType = CommandType.Text;//cmdType;
            if (cmdParms != null)
            {
                foreach (SQLiteParameter parm in cmdParms)
                    cmd.Parameters.Add(parm);
            }
        }

        #endregion

        public static DataTable GetSchema(string collectionName, string[] restictionValues)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                DataTable dt = new DataTable();
                try
                {
                    dt.Clear();
                    connection.Open();
                    dt = connection.GetSchema(collectionName, restictionValues);

                }
                catch
                {
                    dt = null;
                }

                return dt;

            }

        }
        public static DataTable GetSchema(string collectionName)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                DataTable dt = new DataTable();
                try
                {
                    dt.Clear();
                    connection.Open();
                    dt = connection.GetSchema(collectionName);

                }
                catch
                {
                    dt = null;
                }
                return dt;
            }



        }



        public DataTable ReturnDataTable(string SQL, int StartIndex, int PageSize)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                DataTable dt = new DataTable();
                try
                {
                    connection.Open();
                    SQLiteDataAdapter command = new SQLiteDataAdapter(SQL, connection);

                    command.Fill(StartIndex, PageSize, dt);
                }
                catch (System.Data.SQLite.SQLiteException ex)
                {
                    throw new Exception(ex.Message);
                }
                return dt;
            }
        }

        public DataTable ReturnDataTable(string SQL)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                DataTable dt = new DataTable();
                try
                {
                    connection.Open();
                    SQLiteDataAdapter command = new SQLiteDataAdapter(SQL, connection);
                    command.Fill(dt);
                }
                catch (System.Data.SQLite.SQLiteException ex)
                {
                    throw new Exception(ex.Message);
                }
                return dt;
            }
        }

        public int RunProcedure(string storedProcName)
        {
            throw new NotImplementedException();
        }

        public IList<string> GetDataBaseInfo()
        {
            throw new NotImplementedException();
        }

        public IList<string> GetTableInfo()
        {
            IList<string> tableList = new List<string>();
            string[] rs = new string[] { null, null, null, "table" };
            DataTable dt = GetSchema("tables", rs);
            int num = dt.Rows.Count;
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow _DataRowItem in dt.Rows)
                {
                    tableList.Add(_DataRowItem["table_name"].ToString());
                }
            }
            return tableList;
        }

        public IList<string> GetColumnInfo(string TableName)
        {
            string[] restrictions = new string[] { null, null, TableName };
            DataTable tableinfo = GetSchema("Columns", restrictions);
            IList<string> List = new List<string>();
            int count = tableinfo.Rows.Count;
            if (count > 0)
            {

                foreach (DataRow _DataRowItem in tableinfo.Rows)
                {
                    List.Add(_DataRowItem["Column_Name"].ToString());
                }
            }

            return List;

        }

        public IList<string> GetProcInfo()
        {
            IList<string> List = new List<string>();
            List.Add("");
            return List;
        }

        public IList<string> GetViewInfo()
        {
            IList<string> List = new List<string>();
            string[] rs = new string[] { null, null, null, "BASE TABLE" };
            DataTable dt = GetSchema("Views");
            int num = dt.Rows.Count;
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow _DataRowItem in dt.Rows)
                {
                    List.Add(_DataRowItem["table_name"].ToString());
                }
            }
            return List;
        }

        public int ReturnTbCount(string tb_name)
        {
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {

                try
                {
                    string SQL = "select  count(*)   from " + tb_name;
                    connection.Open();
                    SQLiteCommand cmd = new SQLiteCommand(SQL, connection);
                    int count = int.Parse(cmd.ExecuteScalar().ToString());
                    return count;
                }
                catch (System.Data.SQLite.SQLiteException ex)
                {
                    throw new Exception(ex.Message);
                }

            }
        }
        #region 批量导入数据库
        public bool DatatableImport(IList<string> maplist, string TableName, DataTable dt)
        {
            using (SQLiteConnection conn = new SQLiteConnection(connectionString))
            {
                conn.Open();
                SQLiteCommand cmd = new SQLiteCommand();
                cmd.Connection = conn;
                SQLiteTransaction tx = conn.BeginTransaction();
                cmd.Transaction = tx;
                try
                {
                    foreach (DataRow r in dt.Rows)
                    {
                        cmd.CommandText = GenerateInserSql(maplist, TableName, r);
                        cmd.ExecuteNonQuery();
                    }

                    tx.Commit();
                    return true;
                }
                catch (System.Data.SQLite.SQLiteException E)
                {
                    tx.Rollback();
                    throw new Exception(E.Message);
                }
            }


        }


   

        private string GenerateInserSql(IList<string> maplist, string TableName, DataRow row)
        {

            var names = new StringBuilder();
            var values = new StringBuilder();
            bool first = true;
            char quote = '"';

            foreach (string c in maplist)
            {
                if (!first)
                {
                    names.Append(",");
                    values.Append(",");
                }
                names.Append("[" + c + "]");
                values.Append("\"" + row[c].ToString().Replace(quote.ToString(), string.Concat(quote, quote)) + "\"");
                first = false;

            }

            string sql = string.Format("INSERT INTO [{0}] ({1}) VALUES ({2})", TableName, names, values);
            return sql;
        }

        public int BulkCopyFromOpenrowset(IList<string> maplist, string TableName, string filename)
        {
            return -1;
        }

        #endregion
        public int TruncateTable(string TableName)
        {

            return ExecuteSql("delete from [" + TableName+"]");

        }





    }
}
