using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Collections.Generic;
using System.Linq;

namespace DataPie.Core
{
    public class FileToDB
    {

       // Excel2003、Excel2007、ACCESS2007、CSV文件导入
        public static DataTable GetDataTableFromFile(string path, string tname)
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
            DataTable dt = new DataTable(tname);
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
                case ".csv":
                    conn = String.Format(conn, ace, path.Substring(0, path.LastIndexOf('\\')), "text;Excel 12.0", hdr, imex);
                    select = string.Format("SELECT * FROM [{0}]", Path.GetFileName(path));
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


    

       //从文件夹获取csv文件
        public static List<FileInfo> GetFilelist(string DirectoryPath, bool rec)
        {
            List<FileInfo> fileList = FileManager.FileList(DirectoryPath, rec);

            var csvlist = from f in fileList
                          where f.Extension == ".csv"
                          select f;
            return csvlist.ToList();

        }

        public static DataTable GetDataTableFromCSV(FileInfo file)
        {
            string tablename = file.Name.Substring(0, file.Name.LastIndexOf("."));
            return  GetDataTableFromFile(file.FullName, tablename);
        }

 
    }
}
