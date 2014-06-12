using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataPie.Core
{
    public class Common
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
    }
}
