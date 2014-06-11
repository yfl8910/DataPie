using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.IO;
using System.Data;
using System.Diagnostics;

namespace DataPie.Core
{
    class DBToAccess
    {


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
