using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Data.OleDb;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.Windows.Forms;
using Kent.Boogaart.KBCsv;


namespace DataPie
{
    class UiServices
    {

        public static DBConfig db;

        #region Excel2003、Excel2007、ACCESS2007导入操作

        public static DataTable GetExcelDataTable(string path, string tname)
        {
            string ace = "Microsoft.ACE.OLEDB.12.0";
            string jet = "Microsoft.Jet.OLEDB.4.0";
            string xl2007 = "Excel 12.0 Xml";
            string xl2003 = "Excel 8.0";
            string imex = "IMEX=1";
            string hdr = "Yes";
            string conn = "Provider={0};Data Source={1};Extended Properties=\"{2};HDR={3};{4}\";";
            string select = "";
            string ext = Path.GetExtension(path);
            OleDbDataAdapter oda;
            DataTable dt = new DataTable("data");
            switch (ext.ToLower())
            {
                case ".xlsx":
                    conn = String.Format(conn, ace, Path.GetFullPath(path), xl2007, hdr, imex);
                    select = string.Format("SELECT * FROM [{0}$]", tname);
                    break;
                case ".xls":
                    conn = String.Format(conn, jet, Path.GetFullPath(path), xl2003, hdr, imex);
                    select = string.Format("SELECT * FROM [{0}$]", tname);
                    break;
                case ".accdb":
                    conn = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source= {0};Persist Security Info=False;", path);
                    select = string.Format("SELECT * FROM [{0}]", tname);
                    break;
                default:
                    throw new Exception("File Not Supported!");
            }
            OleDbConnection con = new OleDbConnection(conn);
            con.Open();
            oda = new OleDbDataAdapter(select, con);
            oda.Fill(dt);
            con.Close();
            return dt;
        }


        #endregion

        #region Excel保存操作，OpenXML方式

        /// <summary>
        /// 已有工作簿中，添加新的sheet保存
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


        #endregion

        #region 数据库中获取DataTable内存表

      
        /// <summary>
        /// 获取数据表，根据数据库名获取
        /// </summary>
        public static DataTable GetDataTableFromSQL(string sql)
        {
            DataTable dt = new DataTable();
            dt = db.DBProvider.ReturnDataTable(sql);
            return dt;
        }

        /// <summary>
        /// 获取数据表，根据数据库名获取
        /// </summary>
        public static DataTable GetDataTableFromName(string TabalName)
        {
            DataTable dt = new DataTable();
            string sql = "select * from  [" + TabalName + "]";
            dt = db.DBProvider.ReturnDataTable(sql);
            return dt;
        }

   

        /// <summary>
        /// 获取数据表,分页版
        /// </summary>
        public static DataTable GetDBDataTable(string sql, int PageSize, int CurrentPage)
        {
            //int RecordCount
            DataTable dt = new DataTable();
            dt = db.DBProvider.ReturnDataTable(sql, PageSize * (CurrentPage - 1), PageSize);
            return dt;
        }


        #endregion

        #region Excel导出操作

        /// <summary>
        /// 单个数据库表格导出到一个excel工作簿
        /// </summary>
        public static int ExportExcel(string TabelName, string filename)
        {
            //string filename = ShowFileDialog(TabelName);

            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                string sql = "select * from  [" + TabelName + "]";
                DataTable dt = GetDataTableFromSQL(sql);
                SaveExcel(filename, dt, TabelName); watch.Stop();
                //MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);

            }
            return -1;

        }

        /// <summary>
        /// DataTable导出到一个excel工作簿,新建一个名称SheetName的worksheet
        /// </summary>
        public static int ExportExcel(DataTable dt, string SheetName, string filename)
        {
            //string filename = ShowFileDialog(SheetName);

            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                SaveExcel(filename, dt, SheetName);
                watch.Stop();
                //MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;

        }
        /// <summary>
        /// DataTable导出到一个excel工作簿
        /// </summary>
        public static int ExportExcel(DataTable dt, string filename)
        {
            //string filename = ShowFileDialog("Sheet1");

            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                SaveExcel(filename, dt, "Sheet1");
                watch.Stop();
                //MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;

        }

   
        /// <summary>
        /// 单个数据库表格导出到一个excel工作簿
        /// </summary>
        public static int ExportTemplate(string TabelName, string filename)
        {
            //string filename = ShowFileDialog(TabelName);

            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                string sql = "select * from  [" + TabelName + "]" + " where 1=2";
                DataTable dt = GetDataTableFromSQL(sql);
                SaveExcel(filename, dt, TabelName);
                watch.Stop();
                //MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;

        }

        /// <summary>
        ///  DataTable导出到一个excel工作簿
        /// </summary>
        public static int ExportTemplate(DataTable dt, string TabelName, string filename)
        {
            //string filename = ShowFileDialog(TabelName);

            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                SaveExcel(filename, dt, TabelName);
                watch.Stop();
                //MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;

        }



        /// <summary>
        /// 单个数据库表格导出到一个excel工作簿
        /// </summary>
        public static int ExportExcel(string SheetName, string sql, string filename)
        {
            //string filename = ShowFileDialog(SheetName);

            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                DataTable dt = GetDataTableFromSQL(sql);
                SaveExcel(filename, dt, SheetName);
                watch.Stop();
                //MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;


        }

        /// <summary>
        /// 单个数据库表格导出到多个excel工作簿，分页版本
        /// </summary>
        public static int ExportExcel(string TabelName, int PageSize, string filename)
        {
            //string filename = ShowFileDialog(TabelName);

            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                int RecordCount = db.DBProvider.ReturnTbCount(TabelName);
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
                        DataTable dt = db.DBProvider.ReturnDataTable(sql, PageSize * (i - 1), PageSize);
                        SaveExcel(TabelName, dt, package);
                        package.Save();
                    }
                }
                watch.Stop();
                //MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;


        }

        /// <summary>
        /// 多个数据库表格导出到一个excel工作簿
        /// </summary>
        public static int ExportExcel(IList<string> TabelNames, string filename)
        {
            //string filename = ShowFileDialog(TabelNames[0]);

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
                        dt = db.DBProvider.ReturnDataTable(sqls[i]);
                        SaveExcel(SheetNames[i], dt, package);

                    }
                    package.Save();
                }
                watch.Stop();
                //MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }

            return -1;
        }

        /// <summary>
        /// 多个dt导出到一个sheet
        /// </summary>
        public static int ExportExcel(DataTable[] ds, string SheetName,  string filename)
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
                    int num = ds.Count();
                    int pagesize = ds[0].Rows.Count;
                    ExcelWorksheet ws = package.Workbook.Worksheets.Add(SheetName);
                    ws.Cells["A1"].LoadFromDataTable(ds[0], true);
                    for (int i = 1; i < num; i++)
                    {
                        try
                        {

                            ws.Cells["A" + (i * pagesize + 2).ToString()].LoadFromDataTable(ds[i], false);

                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                    package.Save();
                }
                watch.Stop();
                //MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;

        }

        /// <summary>
        /// 多个dt导出多个sheet
        /// </summary>
        public static int ExportExcel(DataTable[] ds, string filename)
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
                    int num = ds.Count();
                    //int pagesize = ds[0].Rows.Count;

                    for (int i = 0; i < num; i++)
                    {
                        try
                        {
                            ExcelWorksheet ws = package.Workbook.Worksheets.Add(ds[i].TableName.ToString());
                            ws.Cells["A1"].LoadFromDataTable(ds[i], true);

                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }
                    package.Save();
                }
                watch.Stop();
                //MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;

        }


        #endregion

        #region 弹出FileDialog对话框

        public static string ShowFileDialog(string FileName)
        {
            System.Windows.Forms.SaveFileDialog saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog1.Filter = "excel2007|*.xlsx";
            saveFileDialog1.FileName = FileName;
            saveFileDialog1.DefaultExt = ".xlsx";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                return saveFileDialog1.FileName.ToString();
            }
            else
            {
                return null;
            }
        }

        #endregion

        #region 从csv文件夹获取DataTable集合

        public static DataTable[] GetDataTableFromCSV(string DirectoryPath,bool rec)
        {
            List<FileInfo> fileList = FileManager.FileList(DirectoryPath,rec);
            List<FileInfo> files = new List<FileInfo>();
            int n = 0;
            foreach (FileInfo f in fileList)
            {
                if (f.Extension == ".csv")
                {
                    n++;
                    files.Add(f);
                }
            }
            if (n > 0)
            {
                DataTable[] dt = new DataTable[n];
                for (int i = 0; i < n; i++)
                {
                    CsvReader r = new CsvReader(files[i].FullName);
                    r.ReadHeaderRecord();
                    System.Data.DataSet ds = new System.Data.DataSet();
                    //string tablename = "tb";
                    string tablename = files[i].Name.Substring(0, files[i].Name.LastIndexOf("."));
                    int num = r.Fill(ds, tablename);
                    dt[i] = ds.Tables[0];
                }
                return dt;
            }
            else
            {
                return null;
            }


        }

        #endregion

        #region 从DataTable导出到csv文件表

        public static char quote = '"';
        public static int SaveCsv(DataTable dt, string filename)
        {
            Stopwatch watch = Stopwatch.StartNew();
            watch.Start();
            StreamWriter sw = new StreamWriter(filename, false, Encoding.GetEncoding("gb2312"));
            for (int k = 0; k < dt.Columns.Count; k++)
            {
                sw.Write(dt.Columns[k].ColumnName.ToString() + ",");
            }
            sw.Write(Environment.NewLine);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (dt.Rows[i][j].ToString().Contains(quote))
                    {
                        string field = dt.Rows[i][j].ToString();
                        field = field.Replace(quote.ToString(), string.Concat(quote, quote));
                        sw.Write(quote.ToString() + field + quote.ToString() + ",");
                    }
                    else
                    {
                        sw.Write(quote.ToString() + dt.Rows[i][j].ToString() + quote.ToString() + ",");
                    }
                }
                sw.Write(Environment.NewLine);//每写一行数据后换行
            }
            sw.Flush();
            sw.Close();//释放资源
            watch.Stop();
            return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
        }

        public static int WriteDataTableToCsv(string TabelName)
        {

            System.Windows.Forms.SaveFileDialog saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog1.Filter = "csv文件|*.csv";
            saveFileDialog1.FileName = TabelName;
            saveFileDialog1.DefaultExt = ".csv";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string filename = saveFileDialog1.FileName.ToString();
                try
                {
                    DataTable dt = UiServices.GetDataTableFromName(TabelName);
                    int i = SaveCsv(dt, filename);       
                    return i;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            return -1;

        }

        public static int WriteDataTableToCsv(string TabelName, string FileName)
        {
            if (FileName != null)
            { 
                try
                {
                    Stopwatch watch = Stopwatch.StartNew();
                    watch.Start();
                    DataTable dt = UiServices.GetDataTableFromName(TabelName);
                     SaveCsv(dt, FileName);
                    watch.Stop();
                    return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            return -1;

        }

        /// <summary>
        /// 单个数据库表格导出到多个CSV，分页版本
        /// </summary>
        public static int WriteDataTableToCsv(string TabelName, int PageSize, string FileName)
        {
            //string filename = ShowFileDialog(TabelName);

            if (FileName != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                int RecordCount = db.DBProvider.ReturnTbCount(TabelName);
                string sql = "select * from  [" + TabelName + "]";
                int Count = (RecordCount - 1) / PageSize + 1;
                FileInfo newFile = new FileInfo(FileName);
               
                for (int i = 1; i <= Count; i++)
                {
                    string s = FileName.Substring(0, FileName.LastIndexOf("."));
                    StringBuilder newfileName = new StringBuilder(s);
                    newfileName.Append(i + ".csv");
                    newFile = new FileInfo(newfileName.ToString());
                    if (newFile.Exists)
                    {
                        newFile.Delete();
                        newFile = new FileInfo(newfileName.ToString());
                    }

                    DataTable dt = db.DBProvider.ReturnDataTable(sql, PageSize * (i - 1), PageSize);
                   SaveCsv(dt, newfileName.ToString());

                }
                watch.Stop();

                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;


        }

        /// <summary>
        /// 多个数据库表格导出到多个csv
        /// </summary>
        public static int ExportMuticsv(IList<string> TabelNames, string filename)
        {
            //string filename = ShowFileDialog(TabelNames[0]);

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


                for (int i = 0; i < sqls.Count; i++)
                {
                    string s = filename.Substring(0, filename.LastIndexOf("\\"));
                    StringBuilder newfileName = new StringBuilder(s);
                    newfileName.Append("\\" + SheetNames[i] + ".csv");
                    FileInfo newFile = new FileInfo(newfileName.ToString());
                    if (newFile.Exists)
                    {
                        newFile.Delete();
                        newFile = new FileInfo(newfileName.ToString());
                    }
                    dt = db.DBProvider.ReturnDataTable(sqls[i]);
                    SaveCsv(dt, newFile.FullName);

                }

                watch.Stop();
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }

            return -1;
        }

     

        public static int WriteCsvFromsql(string sql, string filename)
        {

            try
            {
                DataTable dt = UiServices.GetDataTableFromSQL(sql);
                int i = SaveCsv(dt, filename);
                return i;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return -1;

        }


        #endregion

        #region 读取文件到dt

        public static DataTable GetDataTableFromfile(FileInfo file)
        {


            DataTable dt = new DataTable();
            CsvReader r = new CsvReader(file.FullName);
            r.ReadHeaderRecord();
            System.Data.DataSet ds = new System.Data.DataSet();
            string tablename = file.Name.Substring(0, file.Name.LastIndexOf("."));
            int num = r.Fill(ds, tablename);
            return ds.Tables[0];
        }

        #endregion

        #region 从文件夹获取csv文件

        public static List<FileInfo> GetFilelist(string DirectoryPath, bool rec)
        {
            List<FileInfo> fileList = FileManager.FileList(DirectoryPath, rec);
            List<FileInfo> files = new List<FileInfo>();
            int n = 0;
            foreach (FileInfo f in fileList)
            {
                if (f.Extension == ".csv")
                {
                    n++;
                    files.Add(f);
                }
            }
            if (n > 0)
            {
                return files;
            }
            else
            {
                return null;
            }


        }

        #endregion


    }
}
