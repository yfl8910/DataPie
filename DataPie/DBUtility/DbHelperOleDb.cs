using System;
using System.Collections;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;
using System.Text;

namespace DataPie.DBUtility
{
    /// <summary>
    ///  ACCESS数据库基础访问类
    /// </summary>
    public  class DbHelperOleDb : IDBUtility
    {
        //数据库连接字符串(web.config来配置)，可以动态更改connectionString支持多数据库.		
        public static string connectionString ="";     		
        public DbHelperOleDb()
        {
        }
        public  DbHelperOleDb(string strConnectionString)
        {
            connectionString = strConnectionString;
        }


        #region 公用方法
       
        //public static int GetMaxID(string FieldName, string TableName)
        //{
        //    string strsql = "select max(" + FieldName + ")+1 from " + TableName;
        //    object obj = DbHelperSQL.GetSingle(strsql);
        //    if (obj == null)
        //    {
        //        return 1;
        //    }
        //    else
        //    {
        //        return int.Parse(obj.ToString());
        //    }
        //}
        //public static bool Exists(string strSql)
        //{
        //    object obj = DbHelperSQL.GetSingle(strSql);
        //    int cmdresult;
        //    if ((Object.Equals(obj, null)) || (Object.Equals(obj, System.DBNull.Value)))
        //    {
        //        cmdresult = 0;
        //    }
        //    else
        //    {
        //        cmdresult = int.Parse(obj.ToString());
        //    }
        //    if (cmdresult == 0)
        //    {
        //        return false;
        //    }
        //    else
        //    {
        //        return true;
        //    }
        //}
        public static bool Exists(string strSql, params OleDbParameter[] cmdParms)
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
        public  int ExecuteSql(string SQLString)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                using (OleDbCommand cmd = new OleDbCommand(SQLString, connection))
                {
                    try
                    {
                        connection.Open();
                        cmd.CommandTimeout = 1000;
                        int rows = cmd.ExecuteNonQuery();
                        return rows;
                    }
                    catch (System.Data.OleDb.OleDbException E)
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
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                OleDbCommand cmd = new OleDbCommand();
                cmd.Connection = conn;
                OleDbTransaction tx = conn.BeginTransaction();
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
                catch (System.Data.OleDb.OleDbException E)
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
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand cmd = new OleDbCommand(SQLString, connection);
                System.Data.OleDb.OleDbParameter myParameter = new System.Data.OleDb.OleDbParameter("@content", OleDbType.VarChar);
                myParameter.Value = content;
                cmd.Parameters.Add(myParameter);
                try
                {
                    connection.Open();
                    int rows = cmd.ExecuteNonQuery();
                    return rows;
                }
                catch (System.Data.OleDb.OleDbException E)
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
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand cmd = new OleDbCommand(strSQL, connection);
                System.Data.OleDb.OleDbParameter myParameter = new System.Data.OleDb.OleDbParameter("@fs", OleDbType.Binary);
                myParameter.Value = fs;
                cmd.Parameters.Add(myParameter);
                try
                {
                    connection.Open();
                    int rows = cmd.ExecuteNonQuery();
                    return rows;
                }
                catch (System.Data.OleDb.OleDbException E)
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
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                using (OleDbCommand cmd = new OleDbCommand(SQLString, connection))
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
                    catch (System.Data.OleDb.OleDbException e)
                    {
                        connection.Close();
                        throw new Exception(e.Message);
                    }
                }
            }
        }
        /// <summary>
        /// 执行查询语句，返回OleDbDataReader
        /// </summary>
        /// <param name="strSQL">查询语句</param>
        /// <returns>OleDbDataReader</returns>
        public IDataReader ExecuteReader(string strSQL)
        {
            OleDbConnection connection = new OleDbConnection(connectionString);
            OleDbCommand cmd = new OleDbCommand(strSQL, connection);
            cmd.CommandTimeout = 10000;
            try
            {
                connection.Open();
                OleDbDataReader myReader = cmd.ExecuteReader();
                return myReader;
            }
            catch (System.Data.OleDb.OleDbException e)
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
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                DataSet ds = new DataSet();
                try
                {
                    connection.Open();
                    OleDbDataAdapter command = new OleDbDataAdapter(SQLString, connection);
                    command.Fill(ds, "ds");
                }
                catch (System.Data.OleDb.OleDbException ex)
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
        public static int ExecuteSql(string SQLString, params OleDbParameter[] cmdParms)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                using (OleDbCommand cmd = new OleDbCommand())
                {
                    try
                    {
                        PrepareCommand(cmd, connection, null, SQLString, cmdParms);
                        int rows = cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                        return rows;
                    }
                    catch (System.Data.OleDb.OleDbException E)
                    {
                        throw new Exception(E.Message);
                    }
                }
            }
        }


        /// <summary>
        /// 执行多条SQL语句，实现数据库事务。
        /// </summary>
        /// <param name="SQLStringList">SQL语句的哈希表（key为sql语句，value是该语句的OleDbParameter[]）</param>
        public static void ExecuteSqlTran(Hashtable SQLStringList)
        {
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();
                using (OleDbTransaction trans = conn.BeginTransaction())
                {
                    OleDbCommand cmd = new OleDbCommand();
                    try
                    {
                        //循环
                        foreach (DictionaryEntry myDE in SQLStringList)
                        {
                            string cmdText = myDE.Key.ToString();
                            OleDbParameter[] cmdParms = (OleDbParameter[])myDE.Value;
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
        public static object GetSingle(string SQLString, params OleDbParameter[] cmdParms)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                using (OleDbCommand cmd = new OleDbCommand())
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
                    catch (System.Data.OleDb.OleDbException e)
                    {
                        throw new Exception(e.Message);
                    }
                }
            }
        }

        /// <summary>
        /// 执行查询语句，返回OleDbDataReader
        /// </summary>
        /// <param name="strSQL">查询语句</param>
        /// <returns>OleDbDataReader</returns>
        public static OleDbDataReader ExecuteReader(string SQLString, params OleDbParameter[] cmdParms)
        {
            OleDbConnection connection = new OleDbConnection(connectionString);
            OleDbCommand cmd = new OleDbCommand();
            try
            {
                PrepareCommand(cmd, connection, null, SQLString, cmdParms);
                OleDbDataReader myReader = cmd.ExecuteReader();
                cmd.Parameters.Clear();
                return myReader;
            }
            catch (System.Data.OleDb.OleDbException e)
            {
                throw new Exception(e.Message);
            }

        }

        /// <summary>
        /// 执行查询语句，返回DataSet
        /// </summary>
        /// <param name="SQLString">查询语句</param>
        /// <returns>DataSet</returns>
        public static DataSet Query(string SQLString, params OleDbParameter[] cmdParms)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                OleDbCommand cmd = new OleDbCommand();
                PrepareCommand(cmd, connection, null, SQLString, cmdParms);
                using (OleDbDataAdapter da = new OleDbDataAdapter(cmd))
                {
                    DataSet ds = new DataSet();
                    try
                    {
                        da.Fill(ds, "ds");
                        cmd.Parameters.Clear();
                    }
                    catch (System.Data.OleDb.OleDbException ex)
                    {
                        throw new Exception(ex.Message);
                    }
                    return ds;
                }
            }
        }
        /// <summary>
        /// 运行SQL语句,返回DataTable对象
        /// </summary>
        public DataTable ReturnDataTable(string SQL)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                DataTable dt = new DataTable();
                try
                {
                    connection.Open();
                    OleDbDataAdapter command = new OleDbDataAdapter(SQL, connection);
                    command.Fill(dt);
                }
                catch (System.Data.SqlClient.SqlException ex)
                {
                    throw new Exception(ex.Message);
                }
                return dt;
            }
        }
        /// <summary>
        /// 运行SQL语句,返回DataTable对象
        /// </summary>
        public DataTable ReturnDataTable(string SQL, int StartIndex, int PageSize)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                DataTable dt = new DataTable();
                try
                {
                    connection.Open();
                    OleDbDataAdapter command = new OleDbDataAdapter(SQL, connection);
                    command.Fill(StartIndex, PageSize, dt);
                }
                catch (System.Data.SqlClient.SqlException ex)
                {
                    throw new Exception(ex.Message);
                }
                return dt;
            }
        }



        private static void PrepareCommand(OleDbCommand cmd, OleDbConnection conn, OleDbTransaction trans, string cmdText, OleDbParameter[] cmdParms)
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
                foreach (OleDbParameter parm in cmdParms)
                    cmd.Parameters.Add(parm);
            }
        }

        #endregion

        /// <summary>
        /// 执行存储过程，返回影响行数
        /// </summary>
        public int RunProcedure(string storedProcName)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                int result;
                connection.Open();
                OleDbCommand command = new OleDbCommand(storedProcName,connection);
                command.CommandTimeout = 1000;
                command.CommandType = CommandType.StoredProcedure;
                result = command.ExecuteNonQuery();
                return result;
            }
        }
   

        #region 架构信息
        /// <summary>
        /// 根据条件，返回架构信息	
        /// </summary>
        /// <param name="collectionName">集合名称</param>
        /// <param name="restictionValues">约束条件</param>
        /// <returns>DataTable</returns>
        public static DataTable GetSchema(string collectionName, string[] restictionValues)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
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
        /// <summary>
        /// 返回指定名称的架构信息	
        /// </summary>
        /// <param name="collectionName">集合名称</param>
        /// <returns>DataTable</returns>
        public static DataTable GetSchema(string collectionName)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
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
        public IList<string> GetDataBaseInfo()
        {
            return null;
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
            if (tableinfo.Rows.Count > 0)
            {
                var t = tableinfo.Select(null, "ordinal_position");
                foreach (DataRow _DataRowItem in t)
                {
                    List.Add(_DataRowItem["Column_Name"].ToString());
                }
            }

            return List;

        }

        public bool IF_Proc(string sql)
        {

            if (sql.ToUpper().Contains("DELETE") || sql.ToUpper().Contains("UPDATE"))

                return true;

            else if (sql.ToUpper().Contains("SELECT") && sql.ToUpper().Contains("INTO"))

                return true;

            else return false;

        }
        public IList<string> GetProcInfo()
        {

            IList<string> List = new List<string>();

            DataTable dt = GetSchema("Procedures");

            int num = dt.Rows.Count;

            if (dt != null && dt.Rows.Count > 0)
            {

                foreach (DataRow _DataRowItem in dt.Rows)
                {
                    if (IF_Proc(_DataRowItem["PROCEDURE_DEFINITION"].ToString()))

                    { 
                        List.Add(_DataRowItem["PROCEDURE_NAME"].ToString());
                    }

                }

            }

            return List;

        }
        public IList<string> GetFunctionInfo()
        {
            IList<string> List = new List<string>();
            DataTable dt = GetSchema("Procedures");
            int num = dt.Rows.Count;
            if (dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow _DataRowItem in dt.Rows)
                {
                    if (_DataRowItem["PROCEDURE_TYPE"].ToString().ToUpper() == "FUNCTION")
                    { List.Add(_DataRowItem["PROCEDURE_NAME"].ToString()); }

                }
            }
            return List;
        }
        public IList<string> GetViewInfo()
        {

            IList<string> List = new List<string>();

            string[] rs = new string[] { null, null, null, "BASE TABLE" };

            DataTable dt = GetSchema("views");

            int num = dt.Rows.Count;

            if (dt.Rows.Count > 0)
            {

                foreach (DataRow _DataRowItem in dt.Rows)
                {

                    List.Add(_DataRowItem["table_name"].ToString());

                }

            }

            //添加被架构默认为存储过程的视图

            dt = GetSchema("Procedures");

            num = dt.Rows.Count;

            if (dt != null && dt.Rows.Count > 0)
            {

                foreach (DataRow _DataRowItem in dt.Rows)
                {



                    if (!IF_Proc(_DataRowItem["PROCEDURE_DEFINITION"].ToString()))

                    {
                        List.Add(_DataRowItem["PROCEDURE_NAME"].ToString());
                    }

                }

            }

            return List;

        }
        public int ReturnTbCount(string tb_name)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {

                try
                {
                    string SQL = "select  count(*)   from " + tb_name;
                    connection.Open();
                    OleDbCommand cmd = new OleDbCommand(SQL, connection);
                    int count = int.Parse(cmd.ExecuteScalar().ToString());
                    return count;
                }
                catch (System.Data.SqlClient.SqlException ex)
                {
                    throw new Exception(ex.Message);
                }

            }
        }



        #endregion
        #region 批量导入数据库
        public bool SqlBulkCopyImport(  IList<string> maplist, string TableName, DataTable dt)
        {
            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    OleDbDataAdapter adapter = new OleDbDataAdapter("select * from " + TableName + "  where 1=0", connection);
                    OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter);
                    int rowcount = dt.Rows.Count;
                    for (int n = 0; n < rowcount; n++)
                    {
                        dt.Rows[n].SetAdded();
                    }
                    //adapter.UpdateBatchSize = 1000;
                    adapter.Update(dt);
                }
                return true;
            }
            catch (Exception e)
            {
                throw e;
            }

        }

          public int BulkCopyFromOpenrowset( IList<string> maplist, string TableName, string filename)
        {
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                var names = new StringBuilder();
                bool first = true;


                foreach (string c in maplist)
                {
                    if (!first)
                    {
                        names.Append(",");
                    }
                    names.Append("[" + c + "]");
                    first = false;

                }
          
              string sql = string.Format("insert into  [{0}] ({1}) select {2} from [Excel 12.0;HDR=Yes;IMEX=1;database={3}].[{4}$]", TableName, names, names, filename, TableName);  
                using (OleDbCommand cmd = new OleDbCommand(sql, connection))
                {
                    try
                    {
                        connection.Open();
                        int rows = cmd.ExecuteNonQuery();
                        return rows;
                    }
                    catch (System.Data.OleDb.OleDbException E)
                    {
                        connection.Close();
                        throw new Exception(E.Message);
                    }
                }
            }

        }
          public bool DatatableImport(string constring, IList<string> maplist, string TableName, DataTable dt, bool needCreate)
          {
              using (OleDbConnection conn = new OleDbConnection(constring))
              {
                  conn.Open();
                  OleDbCommand cmd = new OleDbCommand();
                  cmd.Connection = conn;
                  OleDbTransaction tx = conn.BeginTransaction();
                  cmd.Transaction = tx;
                  if (needCreate)
                  {
                      string creatDLL = Core.DBToAccess.CreateTable(dt, TableName);
                      cmd.CommandText = creatDLL;
                      cmd.ExecuteNonQuery();

                  }
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
                  catch (System.Data.OleDb.OleDbException E)
                  {
                      tx.Rollback();
                      throw new Exception(E.Message);
                  }
              }
             
              
          }
          public bool DatatableImport(IList<string> maplist, string TableName, DataTable dt)
          {
              using (OleDbConnection conn = new OleDbConnection(connectionString))
              {
                  conn.Open();
                  OleDbCommand cmd = new OleDbCommand();
                  cmd.Connection = conn;
                  OleDbTransaction tx = conn.BeginTransaction();
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
                  catch (System.Data.OleDb.OleDbException E)
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
       
        #endregion
          public int TruncateTable(string TableName)
        {

            return ExecuteSql("delete from [" + TableName + "]");

        }

    }
}
