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
using Excel = Microsoft.Office.Interop.Excel;

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

        #region Excel保存操作，Microsoft.Office.Interop.Excel方式

        /// <summary>
        ///DataTable导出到一个excel工作簿
        /// </summary>
        public static int ExportOfficeExcel(DataTable dt, string SheetName)
        {
            string saveFile = ShowFileDialog(SheetName);
            if (saveFile != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                Microsoft.Office.Interop.Excel.Application rptExcel = new Microsoft.Office.Interop.Excel.Application();
                if (rptExcel == null)
                {
                    MessageBox.Show("无法打开EXcel，请检查Excel是否可用或者是否安装好Excel", "系统提示");
                    return 0;
                }
                int rowCount = dt.Rows.Count;//行数
                int columnCount = dt.Columns.Count;//列数
                //float percent = 0;//导出进度
                //保存文化环境
                System.Globalization.CultureInfo currentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                Excel.Workbook workbook = rptExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets.get_Item(1);
                worksheet.Name = SheetName;

                //填充列标题
                for (int i = 0; i < columnCount; i++)
                {
                    worksheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
                }

                //创建对象数组存储DataTable的数据，这样的效率比直接将Datateble的数据填充worksheet.Cells[row,col]高
                object[,] objData = new object[rowCount, columnCount];
                if (rowCount > 1)
                {
                    //填充内容到对象数组
                    for (int r = 0; r < rowCount; r++)
                    {
                        for (int col = 0; col < columnCount; col++)
                        {
                            //objData[r, col] = dt.Rows[r][col].ToString();
                            objData[r, col] = dt.Rows[r][col];
                        }
                        //percent = ((float)(r + 1) * 100) / rowCount;
                        System.Windows.Forms.Application.DoEvents();
                    }
                    Excel.Range range = worksheet.get_Range("A2", "A2").get_Resize(rowCount, columnCount);
                    range.NumberFormat = "@";//设置数字文本格式
                    range.Value2 = objData;
                }


                //恢复文化环境
                System.Threading.Thread.CurrentThread.CurrentCulture = currentCI;
                try
                {
                    workbook.Saved = true;
                    workbook.SaveCopyAs(saveFile);//以复制的形式保存在已有的文档里   
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出文件出错，文件可能正被打开，具体原因：" + ex.Message, "出错信息");
                }
                finally
                {
                    dt.Dispose();
                    rptExcel.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rptExcel);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    GC.Collect();
                    KillAllExcel();
                }
                watch.Stop();
                MessageBox.Show("恭喜，数据已经成功导出为Excel文件！", "成功导出");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;
        }

        /// <summary>
        ///根据TableName导出excel工作簿
        /// </summary>
        public static int ExportOfficeExcel(string TableName)
        {
            string saveFile = ShowFileDialog(TableName);
            if (saveFile != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();

                Microsoft.Office.Interop.Excel.Application rptExcel = new Microsoft.Office.Interop.Excel.Application();
                if (rptExcel == null)
                {
                    MessageBox.Show("无法打开EXcel，请检查Excel是否可用或者是否安装好Excel", "系统提示");
                    return -1;
                }
                string sql = "select * from  [" + TableName + "]";
                DataTable dt = GetDataTableFromSQL(sql);
                int rowCount = dt.Rows.Count;//行数
                int columnCount = dt.Columns.Count;//列数
                //float percent = 0;//导出进度
                //保存文化环境
                System.Globalization.CultureInfo currentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                Excel.Workbook workbook = rptExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets.get_Item(1);
                worksheet.Name = TableName;

                //填充列标题
                for (int i = 0; i < columnCount; i++)
                {
                    worksheet.Cells[1, i + 1] = dt.Columns[i].ColumnName;
                }

                //创建对象数组存储DataTable的数据，这样的效率比直接将Datateble的数据填充worksheet.Cells[row,col]高
                object[,] objData = new object[rowCount, columnCount];
                if (rowCount > 1)
                {
                    //填充内容到对象数组
                    for (int r = 0; r < rowCount; r++)
                    {
                        for (int col = 0; col < columnCount; col++)
                        {
                            //objData[r, col] = dt.Rows[r][col].ToString();
                            objData[r, col] = dt.Rows[r][col];
                        }
                        //percent = ((float)(r + 1) * 100) / rowCount;
                        System.Windows.Forms.Application.DoEvents();
                    }
                    Excel.Range range = worksheet.get_Range("A2", "A2").get_Resize(rowCount, columnCount);
                    range.NumberFormat = "@";//设置数字文本格式
                    range.Value2 = objData;
                }


                //恢复文化环境
                System.Threading.Thread.CurrentThread.CurrentCulture = currentCI;
                try
                {
                    workbook.Saved = true;
                    workbook.SaveCopyAs(saveFile);//以复制的形式保存在已有的文档里   
                }
                catch (Exception ex)
                {
                    MessageBox.Show("导出文件出错，文件可能正被打开，具体原因：" + ex.Message, "出错信息");
                }
                finally
                {
                    dt.Dispose();
                    rptExcel.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rptExcel);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    GC.Collect();
                    KillAllExcel();
                }
                watch.Stop();
                MessageBox.Show("恭喜，数据已经成功导出为Excel文件！", "成功导出");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;


        }

        /// <summary>
        ///根据TableName导出excel工作簿
        /// </summary>
        public static int ExportOfficeExcel(string TableName, int PageSize)
        {
            string filename = ShowFileDialog(TableName);

            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                int RecordCount = db.DBProvider.ReturnTbCount(TableName);
                string sql = "select * from  [" + TableName + "]";
                int WorkBookCount = (RecordCount - 1) / PageSize + 1;
                FileInfo newFile = new FileInfo(filename);
                for (int i = 1; i <= WorkBookCount; i++)
                {
                    Microsoft.Office.Interop.Excel.Application rptExcel = new Microsoft.Office.Interop.Excel.Application();
                    if (rptExcel == null)
                    {
                        MessageBox.Show("无法打开EXcel，请检查Excel是否可用或者是否安装好Excel", "系统提示");
                        return -1;
                    }
                    string s = filename.Substring(0, filename.LastIndexOf("."));
                    StringBuilder newfileName = new StringBuilder(s);
                    newfileName.Append(i + ".xlsx");
                    newFile = new FileInfo(newfileName.ToString());
                    if (newFile.Exists)
                    {
                        newFile.Delete();
                        newFile = new FileInfo(newfileName.ToString());
                    }
                    DataTable dt = db.DBProvider.ReturnDataTable(sql, PageSize * (i - 1), PageSize);

                    int rowCount = dt.Rows.Count;//行数
                    int columnCount = dt.Columns.Count;//列数
                    //float percent = 0;//导出进度
                    //保存文化环境
                    System.Globalization.CultureInfo currentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                    Excel.Workbook workbook = rptExcel.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets.get_Item(1);
                    worksheet.Name = TableName;

                    //填充列标题
                    for (int c = 0; c < columnCount; c++)
                    {
                        worksheet.Cells[1, c + 1] = dt.Columns[c].ColumnName;
                    }

                    //创建对象数组存储DataTable的数据，这样的效率比直接将Datateble的数据填充worksheet.Cells[row,col]高
                    object[,] objData = new object[rowCount, columnCount];
                    if (rowCount > 1)
                    {
                        //填充内容到对象数组
                        for (int r = 0; r < rowCount; r++)
                        {
                            for (int col = 0; col < columnCount; col++)
                            {
                                //objData[r, col] = dt.Rows[r][col].ToString();
                                objData[r, col] = dt.Rows[r][col];
                            }
                            //percent = ((float)(r + 1) * 100) / rowCount;
                            System.Windows.Forms.Application.DoEvents();
                        }
                        Excel.Range range = worksheet.get_Range("A2", "A2").get_Resize(rowCount, columnCount);
                        range.NumberFormat = "@";//设置数字文本格式
                        range.Value2 = objData;
                    }


                    //恢复文化环境
                    System.Threading.Thread.CurrentThread.CurrentCulture = currentCI;
                    try
                    {
                        workbook.Saved = true;
                        workbook.SaveCopyAs(newfileName.ToString());//以复制的形式保存在已有的文档里   
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("导出文件出错，文件可能正被打开，具体原因：" + ex.Message, "出错信息");
                    }
                    finally
                    {
                        dt.Dispose();
                        rptExcel.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rptExcel);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);

                    }
                }

                GC.Collect();
                KillAllExcel();
                watch.Stop();
                MessageBox.Show("恭喜，数据已经成功导出为Excel文件！", "成功导出");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;



        }
        /// <summary>
        /// 获得所有的Excel进程
        /// </summary>
        /// <returns>所有的Excel进程</returns>
        private static List<Process> GetExcelProcesses()
        {
            Process[] processes = Process.GetProcesses();
            List<Process> excelProcesses = new List<Process>();
            for (int i = 0; i < processes.Length; i++)
            {
                if (processes[i].ProcessName.ToUpper() == "EXCEL")
                    excelProcesses.Add(processes[i]);
            }
            return excelProcesses;
        }
        private static void KillAllExcel()
        {
            List<Process> excelProcess = GetExcelProcesses();
            for (int i = 0; i < excelProcess.Count; i++)
            {
                excelProcess[i].Kill();
            }
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

        ///// <summary>
        ///// 获取数据表,根据sql获取
        ///// </summary>
        //public static DataTable GetDBDataTable(DBConfig db, string sql)
        //{
        //    DataTable dt = new DataTable();
        //    dt = db.DB.ReturnDataTable(sql);
        //    return dt;
        //}

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

        ///// <summary>
        ///// 获取数据表,分页版
        ///// </summary>
        //public static DataTable GetDBDataTable(DBConfig db, string sql, int PageSize, int CurrentPage)
        //{
        //    //int RecordCount
        //    DataTable dt = new DataTable();
        //    dt = db.DB.ReturnDataTable(sql, PageSize * (CurrentPage - 1), PageSize);
        //    return dt;
        //}

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
        public static int ExportExcel(string TabelName)
        {
            string filename = ShowFileDialog(TabelName);

            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                string sql = "select * from  [" + TabelName + "]";
                DataTable dt = GetDataTableFromSQL(sql);
                SaveExcel(filename, dt, TabelName); watch.Stop();
                MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);

            }
            return -1;

        }

        /// <summary>
        /// DataTable导出到一个excel工作簿,新建一个名称SheetName的worksheet
        /// </summary>
        public static int ExportExcel(DataTable dt, string SheetName)
        {
            string filename = ShowFileDialog(SheetName);

            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                SaveExcel(filename, dt, SheetName);
                watch.Stop();
                MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;

        }
        /// <summary>
        /// DataTable导出到一个excel工作簿
        /// </summary>
        public static int ExportExcel(DataTable dt)
        {
            string filename = ShowFileDialog("Sheet1");

            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                SaveExcel(filename, dt, "Sheet1");
                watch.Stop();
                MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;

        }

        ///// <summary>
        ///// 单个数据库表格导出到一个excel工作簿
        ///// </summary>
        //public static int ExportTemplate(DBConfig db, string TabelName)
        //{
        //    string filename = ShowFileDialog(TabelName);
        //    Stopwatch watch = Stopwatch.StartNew();
        //    watch.Start();
        //    if (filename != null)
        //    {
        //        string sql = "select * from  [" + TabelName + "]" + " where 1=2";
        //        DataTable dt = GetDBDataTable(db, sql);
        //        SaveExcel(filename, dt, TabelName);
        //    }
        //    watch.Stop();
        //    MessageBox.Show("导出成功");
        //    return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);

        //}

        /// <summary>
        /// 单个数据库表格导出到一个excel工作簿
        /// </summary>
        public static int ExportTemplate(string TabelName)
        {
            string filename = ShowFileDialog(TabelName);

            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                string sql = "select * from  [" + TabelName + "]" + " where 1=2";
                DataTable dt = GetDataTableFromSQL(sql);
                SaveExcel(filename, dt, TabelName);
                watch.Stop();
                MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;

        }

        /// <summary>
        ///  DataTable导出到一个excel工作簿
        /// </summary>
        public static int ExportTemplate(DataTable dt, string TabelName)
        {
            string filename = ShowFileDialog(TabelName);

            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                SaveExcel(filename, dt, TabelName);
                watch.Stop();
                MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;

        }



        /// <summary>
        /// 单个数据库表格导出到一个excel工作簿
        /// </summary>
        public static int ExportExcel(string SheetName, string sql)
        {
            string filename = ShowFileDialog(SheetName);

            if (filename != null)
            {
                Stopwatch watch = Stopwatch.StartNew();
                watch.Start();
                DataTable dt = GetDataTableFromSQL(sql);
                SaveExcel(filename, dt, SheetName);
                watch.Stop();
                MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;


        }

        /// <summary>
        /// 单个数据库表格导出到多个excel工作簿，分页版本
        /// </summary>
        public static int ExportExcel(string TabelName, int PageSize)
        {
            string filename = ShowFileDialog(TabelName);

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
                MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }
            return -1;


        }

        /// <summary>
        /// 多个数据库表格导出到一个excel工作簿
        /// </summary>
        public static int ExportExcel(IList<string> TabelNames)
        {
            string filename = ShowFileDialog(TabelNames[0]);

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
                MessageBox.Show("导出成功");
                return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
            }

            return -1;
        }

        /// <summary>
        /// 多个dt出到一个excel工作簿
        /// </summary>
        public static int ExportExcel(DataTable[] ds, string SheetName)
        {
            string filename = ShowFileDialog(SheetName);

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
                MessageBox.Show("导出成功");
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

        public static DataTable[] GetDataTableFromCSV(string DirectoryPath)
        {
            List<FileInfo> fileList = FileManager.FileList(DirectoryPath);
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
                    string tablename = "tb";
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


        public static int WriteDataTableToCsv( string TabelName)
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
                    Stopwatch watch = Stopwatch.StartNew();
                    watch.Start();
                    DataTable dt = UiServices.GetDataTableFromName(TabelName);
                    StreamWriter sw = new StreamWriter(filename, false, Encoding.GetEncoding("gb2312"));
                    StringBuilder sb = new StringBuilder();
                    for (int k = 0; k < dt.Columns.Count; k++)
                    {
                        sb.Append(dt.Columns[k].ColumnName.ToString() + "\t");
                    }
                    sb.Append(Environment.NewLine);

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            //sb.Append(dt.Rows[i][j].ToString().Replace("\t","") + "\t");
                            sb.Append(dt.Rows[i][j].ToString() + "\t");
                        }
                        sb.Append(Environment.NewLine);//每写一行数据后换行
                    }
                    sw.Write(sb.ToString());
                    sw.Flush();
                    sw.Close();//释放资源
                    watch.Stop();
                    MessageBox.Show("导出成功");
                    return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            return -1;

        }








    }
}
