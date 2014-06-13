using System.Data;
using System.Collections.Generic;

namespace DataPie.DBUtility
{
    public interface IDBUtility
    {
        #region 执行SQL操作     
     
        /// <summary>
        /// 运行SQL语句
        /// </summary>
        /// <param name="SQL"></param>
        int ExecuteSql(string SQL);
        int TruncateTable(string TableName);
        IDataReader ExecuteReader(string strSQL);

        #endregion

        #region 返回DataTable对象

        /// <summary>
        /// 运行SQL语句,返回DataTable对象
        /// </summary>
        DataTable ReturnDataTable(string SQL, int StartIndex, int PageSize);
        /// <summary>
        /// 运行SQL语句,返回DataTable对象
        /// </summary>
        DataTable ReturnDataTable(string SQL);
        #endregion

        #region 存储过程操作
        int RunProcedure(string storedProcName);
        #endregion

        #region 获取数据库Schema信息
        /// <summary>
        /// 获取SQL SERVER中数据库列表
        /// </summary>
        IList<string> GetDataBaseInfo();
        IList<string> GetTableInfo();
        IList<string> GetColumnInfo(string TableName);
        IList<string> GetProcInfo();
        //IList<string> GetFunctionInfo();
        IList<string> GetViewInfo();
        int ReturnTbCount(string tb_name);
        #endregion

        #region 批量导入数据库
        /// <summary>
        /// 批量导入数据库
        /// </summary>
        bool DatatableImport(IList<string> maplist, string TableName, DataTable dt);
        int BulkCopyFromOpenrowset(IList<string> maplist, string TableName, string filename);
        #endregion
    }
}
