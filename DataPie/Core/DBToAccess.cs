using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Diagnostics;

namespace DataPie.Core
{
    public class DBToAccess
    {
        // 通过ADOX创建ACCESS数据库文件
        public static void CreatDataBase(string filename)
        {

            if (System.IO.File.Exists(filename))
            {
                System.IO.File.Delete(filename);
            }
            ADOX.Catalog cat = new ADOX.Catalog();
            cat.Create("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filename + ";");
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cat.ActiveConnection);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cat);
        }

        public static string CreateTable(System.Data.DataTable dt, string tabName)
        {
            StringBuilder sb = new StringBuilder();
            int cols = dt.Columns.Count;
            sb.Append("CREATE TABLE [" + tabName + "] (");
            for (int i = 0; i < cols; i++)
            {
                if (i == 0)
                {
                    sb.Append("[" + dt.Columns[i].ColumnName.ToString() + "]" + "  " + MapType(dt.Columns[i].DataType.ToString()));
                }
                else
                {
                    sb.Append(", " + "[" + dt.Columns[i].ColumnName.ToString() + "]" + "  " + MapType(dt.Columns[i].DataType.ToString()));
                }
            }
            sb.Append(" )");
            return sb.ToString();

        }

        // 把.net数据类型转换为Db数据类型
        public static String MapType(String DataType)
        {
            String reType = String.Empty;
            if (DataType.ToString() == "System.String")
            {
                reType = "varchar";
            }
            else if (DataType.ToString() == "System.Decimal" || DataType.ToString() == "System.Double")
            {
                reType = "number";
            }
            else if (DataType.ToString() == "System.DateTime")
            {
                reType = "datetime";

            }
            else
            {
                reType = "int";

            }
            return reType;

        }

        public static int DataTableExportToAccess(DataTable dt, string filename, String tabName)
        {
            Stopwatch watch = Stopwatch.StartNew();
            watch.Start();

            if (dt.Rows.Count <= 0)
            {
                return -1;
            }

            string connString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}", filename);
            DBUtility.DbHelperOleDb oledb = new DBUtility.DbHelperOleDb();

            IList<string> maplist = new List<string>();
            int cols = dt.Columns.Count;

            for (int i = 0; i < cols; i++)
            {
                maplist.Add(dt.Columns[i].ColumnName.ToString());
            }

            oledb.DatatableImport(connString,maplist, tabName, dt,true);

            watch.Stop();
            return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);

        }


    }
}
