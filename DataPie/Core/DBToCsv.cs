using System;
using System.Linq;
using System.Text;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;

namespace DataPie.Core
{
    public class DBToCsv
    {
        private static void WriteHeader(IDataReader reader, StreamWriter sw)
        {
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (i > 0)
                    sw.Write(',');
                sw.Write(reader.GetName(i));
            }
            sw.Write(Environment.NewLine);
        }
        private static void WriteContent(IDataReader reader, StreamWriter sw)
        {
            for (int i = 0; i < reader.FieldCount; i++)
            {
                if (i > 0)
                    sw.Write(',');
                String v = reader[i].ToString();
                if (v.Contains(',') || v.Contains('\n') || v.Contains('\r') || v.Contains('"'))
                {
                    sw.Write('"');
                    sw.Write(v.Replace("\"", "\"\""));
                    sw.Write('"');
                }
                else
                {
                    sw.Write(v);
                }
            }
            sw.Write(Environment.NewLine);
        }

        public static int SaveCsv(IDataReader reader, string filename)
        {
            Stopwatch watch = Stopwatch.StartNew();
            watch.Start();

            StreamWriter sw = new StreamWriter(filename, false, Encoding.GetEncoding("gb2312"));
            WriteHeader(reader, sw);

            while (reader.Read())
            {
                WriteContent(reader, sw);
            }

            sw.Flush();
            sw.Close();//释放资源

            watch.Stop();
            return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);
        }
        public static async Task<int> ExportCsvAsync(IDataReader reader, string filename)
        {
            return await Task.Run( () => { return SaveCsv( reader,  filename);} );
        }

        public static StreamWriter GetStreamWriter(string filename, int outCount)
        {
            string s = filename.Substring(0, filename.LastIndexOf("."));
            StringBuilder newfileName = new StringBuilder(s);
            newfileName.Append(outCount + 1 + ".csv");
            FileInfo newFile = new FileInfo(newfileName.ToString());
            if (newFile.Exists)
            {
                newFile.Delete();
                newFile = new FileInfo(newfileName.ToString());
            }
            StreamWriter sw = new StreamWriter(newfileName.ToString(), false, Encoding.GetEncoding("gb2312"));
            return sw;
        }
    
        public static int SaveCsv(IDataReader reader, string filename, int pagesize)
        {
            int innerCount = 0, outCount = 0;
            Stopwatch watch = Stopwatch.StartNew();
            watch.Start();
            StreamWriter sw = GetStreamWriter(filename, outCount);
            WriteHeader(reader, sw);
            while (reader.Read())
            {
                if (innerCount < pagesize)
                {
                    WriteContent(reader, sw);
                    innerCount++;
                }
                else
                {
                    innerCount = 0;
                    outCount++;
                    sw.Flush();
                    sw.Close();
                    sw = GetStreamWriter(filename, outCount);
                    WriteHeader(reader, sw);
                    WriteContent(reader, sw);
                    innerCount++;
                }
            }
            sw.Flush();
            sw.Close();
            watch.Stop();
            return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);

        }
        public static async Task<int> ExportCsvAsync(IDataReader reader, string filename, int pagesize)
        {
            return await Task.Run(() => { return SaveCsv(reader, filename,pagesize); });
        }

    }
}
