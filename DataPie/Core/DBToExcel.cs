using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;
using System.Data;
using System.Diagnostics;

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








    }
}
