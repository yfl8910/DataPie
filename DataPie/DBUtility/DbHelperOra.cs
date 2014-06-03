using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
//using System.Data.OracleClient;
using System.Configuration;
using Oracle.DataAccess.Client;

namespace DataPie.DBUtility
{
    /// <summary>
    /// 
    /// </summary>
    public class DbHelperOra : IDBUtility
    {
        //数据库连接字符串(web.config来配置)，可以动态更改connectionString支持多数据库.		
        //public static string connectionString = "User Id=;Password=;Data Source=ORCL";   
        public static string connectionString = null;
        public DbHelperOra()
        {
        }
        public DbHelperOra(string strConnectionString)
        {
            connectionString = strConnectionString;
        }

        #region 公用方法

        public static string GetUser_ID()
        {
            int i = connectionString.IndexOf("=");
            int j = connectionString.IndexOf(";");


            return connectionString.Substring(i + 1, j - i - 1).ToUpper();
        }

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

        public static bool Exists(string strSql, params OracleParameter[] cmdParms)
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
        //public static int ExecuteSql(string SQLString)
        //{
        //    using (OracleConnection connection = new OracleConnection(connectionString))
        //    {				
        //        using (OracleCommand cmd = new OracleCommand(SQLString,connection))
        //        {
        //            try
        //            {		
        //                connection.Open();
        //                int rows=cmd.ExecuteNonQuery();
        //                return rows;
        //            }
        //            catch(System.Data.OracleClient.OracleException E)
        //            {					
        //                connection.Close();
        //                throw new Exception(E.Message);
        //            }
        //        }				
        //    }
        //}

        /// <summary>
        /// 执行多条SQL语句，实现数据库事务。
        /// </summary>
        /// <param name="SQLStringList">多条SQL语句</param>		
        public static void ExecuteSqlTran(ArrayList SQLStringList)
        {
            using (OracleConnection conn = new OracleConnection(connectionString))
            {
                conn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = conn;
                OracleTransaction tx = conn.BeginTransaction();
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
                catch (OracleException E)
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
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                OracleCommand cmd = new OracleCommand(SQLString, connection);

                OracleParameter myParameter = new OracleParameter("@content", OracleDbType.NVarchar2);
                myParameter.Value = content;
                cmd.Parameters.Add(myParameter);
                try
                {
                    connection.Open();
                    int rows = cmd.ExecuteNonQuery();
                    return rows;
                }
                catch (OracleException E)
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
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                OracleCommand cmd = new OracleCommand(strSQL, connection);
                OracleParameter myParameter = new OracleParameter("@fs", OracleDbType.LongRaw);
                myParameter.Value = fs;
                cmd.Parameters.Add(myParameter);
                try
                {
                    connection.Open();
                    int rows = cmd.ExecuteNonQuery();
                    return rows;
                }
                catch (OracleException E)
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
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                using (OracleCommand cmd = new OracleCommand(SQLString, connection))
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
                    catch (OracleException e)
                    {
                        connection.Close();
                        throw new Exception(e.Message);
                    }
                }
            }
        }
        /// <summary>
        /// 执行查询语句，返回OracleDataReader ( 注意：调用该方法后，一定要对SqlDataReader进行Close )
        /// </summary>
        /// <param name="strSQL">查询语句</param>
        /// <returns>OracleDataReader</returns>
        public static OracleDataReader ExecuteReader(string strSQL)
        {
            OracleConnection connection = new OracleConnection(connectionString);
            OracleCommand cmd = new OracleCommand(strSQL, connection);
            try
            {
                connection.Open();
                OracleDataReader myReader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                return myReader;
            }
            catch (OracleException e)
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
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                DataSet ds = new DataSet();
                try
                {
                    connection.Open();
                    OracleDataAdapter command = new OracleDataAdapter(SQLString, connection);
                    command.Fill(ds, "ds");
                }
                catch (OracleException ex)
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
        public static int ExecuteSql(string SQLString, params OracleParameter[] cmdParms)
        {
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                using (OracleCommand cmd = new OracleCommand())
                {
                    try
                    {
                        PrepareCommand(cmd, connection, null, SQLString, cmdParms);
                        int rows = cmd.ExecuteNonQuery();
                        cmd.Parameters.Clear();
                        return rows;
                    }
                    catch (OracleException E)
                    {
                        throw new Exception(E.Message);
                    }
                }
            }
        }


        /// <summary>
        /// 执行多条SQL语句，实现数据库事务。
        /// </summary>
        /// <param name="SQLStringList">SQL语句的哈希表（key为sql语句，value是该语句的OracleParameter[]）</param>
        public static void ExecuteSqlTran(Hashtable SQLStringList)
        {
            using (OracleConnection conn = new OracleConnection(connectionString))
            {
                conn.Open();
                using (OracleTransaction trans = conn.BeginTransaction())
                {
                    OracleCommand cmd = new OracleCommand();
                    try
                    {
                        //循环
                        foreach (DictionaryEntry myDE in SQLStringList)
                        {
                            string cmdText = myDE.Key.ToString();
                            OracleParameter[] cmdParms = (OracleParameter[])myDE.Value;
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
        public static object GetSingle(string SQLString, params OracleParameter[] cmdParms)
        {
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                using (OracleCommand cmd = new OracleCommand())
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
                    catch (OracleException e)
                    {
                        throw new Exception(e.Message);
                    }
                }
            }
        }

        /// <summary>
        /// 执行查询语句，返回OracleDataReader ( 注意：调用该方法后，一定要对SqlDataReader进行Close )
        /// </summary>
        /// <param name="strSQL">查询语句</param>
        /// <returns>OracleDataReader</returns>
        public static OracleDataReader ExecuteReader(string SQLString, params OracleParameter[] cmdParms)
        {
            OracleConnection connection = new OracleConnection(connectionString);
            OracleCommand cmd = new OracleCommand();
            try
            {
                PrepareCommand(cmd, connection, null, SQLString, cmdParms);
                OracleDataReader myReader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                cmd.Parameters.Clear();
                return myReader;
            }
            catch (OracleException e)
            {
                throw new Exception(e.Message);
            }

        }

        /// <summary>
        /// 执行查询语句，返回DataSet
        /// </summary>
        /// <param name="SQLString">查询语句</param>
        /// <returns>DataSet</returns>
        public static DataSet Query(string SQLString, params OracleParameter[] cmdParms)
        {
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                OracleCommand cmd = new OracleCommand();
                PrepareCommand(cmd, connection, null, SQLString, cmdParms);
                using (OracleDataAdapter da = new OracleDataAdapter(cmd))
                {
                    DataSet ds = new DataSet();
                    try
                    {
                        da.Fill(ds, "ds");
                        cmd.Parameters.Clear();
                    }
                    catch (OracleException ex)
                    {
                        throw new Exception(ex.Message);
                    }
                    return ds;
                }
            }
        }


        private static void PrepareCommand(OracleCommand cmd, OracleConnection conn, OracleTransaction trans, string cmdText, OracleParameter[] cmdParms)
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
                foreach (OracleParameter parm in cmdParms)
                    cmd.Parameters.Add(parm);
            }
        }

        #endregion

        #region 存储过程操作

        /// <summary>
        /// 执行存储过程 返回SqlDataReader ( 注意：调用该方法后，一定要对SqlDataReader进行Close )
        /// </summary>
        /// <param name="storedProcName">存储过程名</param>
        /// <param name="parameters">存储过程参数</param>
        /// <returns>OracleDataReader</returns>
        public static OracleDataReader RunProcedure(string storedProcName, IDataParameter[] parameters)
        {
            OracleConnection connection = new OracleConnection(connectionString);
            OracleDataReader returnReader;
            connection.Open();
            OracleCommand command = BuildQueryCommand(connection, storedProcName, parameters);
            command.CommandType = CommandType.StoredProcedure;
            returnReader = command.ExecuteReader(CommandBehavior.CloseConnection);
            return returnReader;
        }


        /// <summary>
        /// 执行存储过程
        /// </summary>
        /// <param name="storedProcName">存储过程名</param>
        /// <param name="parameters">存储过程参数</param>
        /// <param name="tableName">DataSet结果中的表名</param>
        /// <returns>DataSet</returns>
        public static DataSet RunProcedure(string storedProcName, IDataParameter[] parameters, string tableName)
        {
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                DataSet dataSet = new DataSet();
                connection.Open();
                OracleDataAdapter sqlDA = new OracleDataAdapter();
                sqlDA.SelectCommand = BuildQueryCommand(connection, storedProcName, parameters);
                sqlDA.Fill(dataSet, tableName);
                connection.Close();
                return dataSet;
            }
        }


        /// <summary>
        /// 构建 OracleCommand 对象(用来返回一个结果集，而不是一个整数值)
        /// </summary>
        /// <param name="connection">数据库连接</param>
        /// <param name="storedProcName">存储过程名</param>
        /// <param name="parameters">存储过程参数</param>
        /// <returns>OracleCommand</returns>
        private static OracleCommand BuildQueryCommand(OracleConnection connection, string storedProcName, IDataParameter[] parameters)
        {
            OracleCommand command = new OracleCommand(storedProcName, connection);
            command.CommandType = CommandType.StoredProcedure;
            if (parameters != null)
            {
                foreach (OracleParameter parameter in parameters)
                {
                    command.Parameters.Add(parameter);
                }
            }

            return command;
        }

        /// <summary>
        /// 执行存储过程，返回影响的行数		
        /// </summary>
        /// <param name="storedProcName">存储过程名</param>
        /// <param name="parameters">存储过程参数</param>
        /// <param name="rowsAffected">影响的行数</param>
        /// <returns></returns>
        public static int RunProcedure(string storedProcName, IDataParameter[] parameters, out int rowsAffected)
        {
            using (OracleConnection connection = new OracleConnection(connectionString))
            {
                int result;
                connection.Open();
                OracleCommand command = BuildIntCommand(connection, storedProcName, parameters);
                rowsAffected = command.ExecuteNonQuery();
                result = (int)command.Parameters["ReturnValue"].Value;
                //Connection.Close();
                return result;
            }
        }

        /// <summary>
        /// 创建 OracleCommand 对象实例(用来返回一个整数值)	
        /// </summary>
        /// <param name="storedProcName">存储过程名</param>
        /// <param name="parameters">存储过程参数</param>
        /// <returns>OracleCommand 对象实例</returns>
        private static OracleCommand BuildIntCommand(OracleConnection connection, string storedProcName, IDataParameter[] parameters)
        {
            OracleCommand command = BuildQueryCommand(connection, storedProcName, parameters);
            command.Parameters.Add(new OracleParameter("ReturnValue",
                 OracleDbType.Int32, 4, ParameterDirection.ReturnValue,
                false, 0, 0, string.Empty, DataRowVersion.Default, null));
            return command;

        }
        #endregion

        public int ExecuteSql(string SQL)
        {

            using (OracleConnection connection = new OracleConnection(connectionString))
            {

                using (OracleCommand cmd = new OracleCommand(SQL, connection))
                {

                    try
                    {

                        cmd.CommandTimeout = 1000;

                        connection.Open();

                        int rows = cmd.ExecuteNonQuery();

                        return rows;

                    }

                    catch (Exception E)
                    {

                        connection.Close();

                        throw new Exception(E.Message);

                    }

                }

            }

        }



        public DataTable ReturnDataTable(string SQL, int StartIndex, int PageSize)
        {

            using (OracleConnection connection = new OracleConnection(connectionString))
            {

                DataTable dt = new DataTable();

                try
                {

                    connection.Open();

                    OracleDataAdapter command = new OracleDataAdapter(SQL, connection);

                    command.Fill(StartIndex, PageSize, dt);

                }

                catch (OracleException ex)
                {

                    throw new Exception(ex.Message);

                }

                return dt;

            }



        }



        public DataTable ReturnDataTable(string SQL)
        {

            using (OracleConnection connection = new OracleConnection(connectionString))
            {

                DataTable dt = new DataTable();

                try
                {

                    connection.Open();

                    OracleDataAdapter command = new OracleDataAdapter(SQL, connection);

                    command.Fill(dt);

                }

                catch (OracleException ex)
                {

                    throw new Exception(ex.Message);

                }

                return dt;

            }

        }



        public int RunProcedure(string storedProcName)
        {

            using (OracleConnection connection = new OracleConnection(connectionString))
            {

                int result;

                connection.Open();

                OracleCommand command = BuildIntCommand(connection, storedProcName, null);

                command.CommandTimeout = 1000;

                result = command.ExecuteNonQuery();

                //result = (int)command.Parameters["ReturnValue"].Value;

                //Connection.Close();

                return result;

            }

        }



        /// <summary>

        /// 根据条件，返回架构信息 

        /// </summary>

        /// <param name="collectionName">集合名称</param>

        /// <param name="restictionValues">约束条件</param>

        /// <returns>DataTable</returns>

        public static DataTable GetSchema(string collectionName, string[] restictionValues)
        {

            using (OracleConnection connection = new OracleConnection(connectionString))
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

            using (OracleConnection connection = new OracleConnection(connectionString))
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



        public System.Collections.Generic.IList<string> GetDataBaseInfo()
        {

            return null;

        }



        public System.Collections.Generic.IList<string> GetTableInfo()
        {

            IList<string> tableList = new List<string>();

            //string[] rs = new string[] { null, null, "User" };

            //DataTable dt = GetSchema("tables", rs);
            DataTable dt = GetSchema("tables");
            int num = dt.Rows.Count;
            string owner = GetUser_ID();
            if (dt.Rows.Count > 0)
            {

                foreach (DataRow _DataRowItem in dt.Rows)
                {
                    if (_DataRowItem["owner"].ToString() == owner)
                    { tableList.Add(_DataRowItem["table_name"].ToString()); }


                }

            }

            return tableList;

        }



        public System.Collections.Generic.IList<string> GetColumnInfo(string TableName)
        {
            string owner = GetUser_ID();
            string[] restrictions = new string[] { owner, TableName };

            DataTable tableinfo = GetSchema("Columns", restrictions);

            IList<string> List = new List<string>();

            int count = tableinfo.Rows.Count;

            if (count > 0)
            {

                //for (int i = 0; i < count; i++)

                //{

                //    List.Add(tableinfo.Rows[i]["Column_Name"].ToString()); 

                //}



                foreach (DataRow _DataRowItem in tableinfo.Rows)
                {

                    List.Add(_DataRowItem["Column_Name"].ToString());

                }

            }



            return List;

        }



        public System.Collections.Generic.IList<string> GetProcInfo()
        {

            IList<string> List = new List<string>();

            DataTable dt = GetSchema("Procedures");
            string owner = GetUser_ID();
            int num = dt.Rows.Count;

            if (dt != null && dt.Rows.Count > 0)
            {

                foreach (DataRow _DataRowItem in dt.Rows)
                {


                    if (_DataRowItem["owner"].ToString() == owner)
                    { List.Add(_DataRowItem["OBJECT_NAME"].ToString()); }




                }

            }

            return List;

        }



        //public System.Collections.Generic.IList<string> GetFunctionInfo()
        //{

        //    IList<string> List = new List<string>();

        //    DataTable dt = GetSchema("Procedures");

        //    int num = dt.Rows.Count;

        //    if (dt != null && dt.Rows.Count > 0)
        //    {

        //        foreach (DataRow _DataRowItem in dt.Rows)
        //        {

        //            if (_DataRowItem["routine_type"].ToString().ToUpper() == "FUNCTION")

        //            { List.Add(_DataRowItem["routine_name"].ToString()); }



        //        }

        //    }

        //    return List;

        //}



        public System.Collections.Generic.IList<string> GetViewInfo()
        {

            IList<string> List = new List<string>();
            string owner = GetUser_ID();
            string[] rs = new string[] { owner };

            DataTable dt = GetSchema("views", rs);

            int num = dt.Rows.Count;

            if (dt.Rows.Count > 0)
            {

                foreach (DataRow _DataRowItem in dt.Rows)
                {

                    List.Add(_DataRowItem["view_name"].ToString());

                }

            }

            return List;

        }



        public int ReturnTbCount(string tb_name)
        {

            using (OracleConnection connection = new OracleConnection(connectionString))
            {



                try
                {

                    string SQL = "select  count(*)   from " + tb_name;

                    connection.Open();

                    OracleCommand cmd = new OracleCommand(SQL, connection);

                    int count = int.Parse(cmd.ExecuteScalar().ToString());

                    return count;

                }

                catch (System.Data.SqlClient.SqlException ex)
                {

                    throw new Exception(ex.Message);

                }



            }

        }


        #region 批量导入数据库
        public bool SqlBulkCopyImport(System.Collections.Generic.IList<string> maplist, string TableName, DataTable dt)
        {

            using (OracleConnection connection = new OracleConnection(connectionString))
            {

                connection.Open();

                using (OracleBulkCopy bulkCopy = new OracleBulkCopy(connection))
                {

                    bulkCopy.DestinationTableName = TableName;

                    foreach (string a in maplist)
                    {

                        bulkCopy.ColumnMappings.Add(a, a);

                    }

                    try
                    {

                        bulkCopy.WriteToServer(dt);

                        return true;

                    }

                    catch (Exception e)
                    {

                        throw e;

                    }
                }
            }
        }

        public int BulkCopyFromOpenrowset(IList<string> maplist, string TableName, string filename)
        { 
            return -1; 
        }
        #endregion
        public int TruncateTable(string TableName)
        {

            return ExecuteSql("TRUNCATE TABLE   " + TableName);

        }

    }
}
