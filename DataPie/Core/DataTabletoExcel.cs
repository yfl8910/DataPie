using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.IO;
using OfficeOpenXml;
using System.Diagnostics;


namespace DataPie.Core
{
    public class DataTableToExcel
    {
       
        public static DataTable GetDataTableFromSQL(string sql)
        {
            DataTable dt = new DataTable();
            dt = DBConfig.db.DBProvider.ReturnDataTable(sql);
            return dt;
        }

        public static DataTable GetDataTableFromName(string TabalName)
        {
            DataTable dt = new DataTable();
            string sql = "select * from  [" + TabalName + "]";
            dt = DBConfig.db.DBProvider.ReturnDataTable(sql);
            return dt;
        }

        /// <summary>
        /// 已有工作簿中，添加新的sheet并保存
        /// </summary>
        public static bool SaveExcel(string SheetName, DataTable dt, ExcelPackage package)
        {

            try
            {
                ExcelWorksheet ws = package.Workbook.Worksheets.Add(SheetName);
                ws.Cells["A1"].LoadFromDataTable(dt, true);

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 保存excel文件，覆盖相同文件名的文件
        /// </summary>
        public static void SaveExcel(string FileName, DataTable dt, string NewSheetName)
        {
            FileInfo newFile = new FileInfo(FileName);
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(FileName);
            }
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                try
                {
                    ExcelWorksheet ws = package.Workbook.Worksheets.Add(NewSheetName);
                    ws.Cells["A1"].LoadFromDataTable(dt, true);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                package.Save();
            }
        }


        /// <summary>
        /// 单表格导出到excel工作簿
        /// </summary>
        public static int ExportTemplate(string TabelName, string filename)
        {
            
            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                string sql = "select * from  [" + TabelName + "]" + " where 1=2";
                DataTable dt = GetDataTableFromSQL(sql);
                SaveExcel(filename, dt, TabelName);
                watch.Stop(); 
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;

        }

        /// <summary>
        /// 单表格导出到一个excel工作簿
        /// </summary>
        public static int ExportExcel(string SheetName, string sql, string filename)
        {     
            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                DataTable dt = GetDataTableFromSQL(sql);
                SaveExcel(filename, dt, SheetName);
                watch.Stop();
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;


        }

        /// <summary>
        /// 单表格导出到多excel工作簿，分页版本
        /// </summary>
        public static int ExportExcel(string TabelName, int PageSize, string filename)
        {
            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                int RecordCount = DBConfig.db.DBProvider.ReturnTbCount(TabelName);
                string sql = "select * from  [" + TabelName + "]";
                int WorkBookCount = (RecordCount - 1) / PageSize + 1;
                FileInfo newFile = new FileInfo(filename);
                for (int i = 1; i <= WorkBookCount; i++)
                {
                    string s = filename.Substring(0, filename.LastIndexOf("."));
                    StringBuilder newfileName = new StringBuilder(s);
                    newfileName.Append(i + ".xlsx");
                    newFile = new FileInfo(newfileName.ToString());
                    if (newFile.Exists)
                    {
                        newFile.Delete();
                        newFile = new FileInfo(newfileName.ToString());
                    }
                    using (ExcelPackage package = new ExcelPackage(newFile))
                    {
                        DataTable dt = DBConfig.db.DBProvider.ReturnDataTable(sql, PageSize * (i - 1), PageSize);
                        SaveExcel(TabelName, dt, package);
                        package.Save();
                    }
                }
                watch.Stop();
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;
        }

        /// <summary>
        /// 多表格导出到一个excel工作簿
        /// </summary>
        public static int ExportExcel(IList<string> TabelNames, string filename)
        {
            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                IList<string> sqls = new List<string>();
                IList<string> SheetNames = new List<string>();
                foreach (var item in TabelNames)
                {
                    SheetNames.Add(item.ToString());
                    sqls.Add("select * from  [" + item.ToString() + "]");
                }
                DataTable dt = new DataTable();

                FileInfo newFile = new FileInfo(filename);
                if (newFile.Exists)
                {
                    newFile.Delete();
                    newFile = new FileInfo(filename);
                }
                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    for (int i = 0; i < sqls.Count; i++)
                    {
                        dt = DBConfig.db.DBProvider.ReturnDataTable(sqls[i]);
                        SaveExcel(SheetNames[i], dt, package);

                    }
                    package.Save();
                }
                watch.Stop();
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }

            return -1;
        }


    }
}
