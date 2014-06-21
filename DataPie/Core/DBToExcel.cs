using System;
using System.IO;
using OfficeOpenXml;
using System.Data;
using System.Diagnostics;
using System.Text;
namespace DataPie.Core
{
    public class DBToExcel
    {

        /// <summary>
        /// 保存excel文件，覆盖相同文件名的文件
        /// </summary>
        public static int SaveExcel(string FileName, string sql, string SheetName)
        {

            Stopwatch watch = Stopwatch.StartNew();
            watch.Start();

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
                    ExcelWorksheet ws = package.Workbook.Worksheets.Add(SheetName);

                    IDataReader reader = DBConfig.db.DBProvider.ExecuteReader(sql);
                    ws.Cells["A1"].LoadFromDataReader(reader, true);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                package.Save();
            }
            watch.Stop();
            return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
        }

        /// <summary>
        /// 工作簿中添加新的SHEET
        /// </summary>
        public static bool SaveExcel(ExcelPackage package, string sql, string SheetName)
        {

            try
            {
                ExcelWorksheet ws = package.Workbook.Worksheets.Add(SheetName);
                IDataReader reader = DBConfig.db.DBProvider.ExecuteReader(sql);
                ws.Cells["A1"].LoadFromDataReader(reader, true);

                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }


        /// <summary>
        /// 多表格导出到一个EXCEL工作簿
        /// </summary>
        public static int ExportExcel(string[] TabelNameArray, string filename)
        {
            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();

                FileInfo newFile = new FileInfo(filename);
                if (newFile.Exists)
                {
                    newFile.Delete();
                    newFile = new FileInfo(filename);
                }


                using (ExcelPackage package = new ExcelPackage(newFile))
                {
                    for (int i = 0; i < TabelNameArray.Length; i++)
                    {
                        string sql = "select * from  [" + TabelNameArray[i] + "]";
                        IDataReader reader = DBConfig.db.DBProvider.ExecuteReader(sql);
                        SaveExcel(package, sql, TabelNameArray[i]);

                    }
                    package.Save();
                }
                watch.Stop();
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }

            return -1;
        }

        /// <summary>
        /// 分拆导出
        /// </summary>
        public static int ExportExcel(string[] TabelNameArray, string filename, string[] whereSQLArr)
        {
            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                FileInfo file = new FileInfo(filename);
                int wherecount = whereSQLArr.Length;
                int count = TabelNameArray.Length;
                for (int i = 0; i < wherecount; i++)
                {
                    string s = filename.Substring(0, filename.LastIndexOf("\\"));
                    StringBuilder newfileName = new StringBuilder(s);
                    int index = whereSQLArr[i].LastIndexOf("=");
                    string sp = whereSQLArr[i].Substring(index + 2, whereSQLArr[i].Length - index - 3);
                    newfileName.Append("\\" + file.Name.Substring(0, file.Name.LastIndexOf(".")) + "_" + sp + ".xlsx");
                    FileInfo newFile = new FileInfo(newfileName.ToString());
                    if (newFile.Exists)
                    {
                        newFile.Delete();
                        newFile = new FileInfo(newfileName.ToString());
                    }

                    for (int j = 0; j < count; j++)
                    {
                        string sql = "select * from [" + TabelNameArray[j] + "]" + whereSQLArr[i];
                        IDataReader reader = DBConfig.db.DBProvider.ExecuteReader(sql);
                        SaveExcel(newfileName.ToString(), sql, TabelNameArray[j]);

                    }


                }

                watch.Stop();
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }

            return -1;
        }




    }
}
