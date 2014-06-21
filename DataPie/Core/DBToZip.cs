using System;
using System.Linq;
using System.Text;
using System.IO.Compression;
using System.Data;
using System.Diagnostics;
using System.IO;

namespace DataPie.Core
{
   public class DBToZip
    {


        public static int DataReaderToZip(String zipFileName, IDataReader reader, string tablename)
        {
            Stopwatch watch = Stopwatch.StartNew();
            watch.Start();
            using (FileStream fsOutput = new FileStream(zipFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                using (ZipArchive archive = new ZipArchive(fsOutput, ZipArchiveMode.Update))
                {
                    ZipArchiveEntry readmeEntry = archive.CreateEntry(tablename + ".csv");
                    using (StreamWriter writer = new StreamWriter(readmeEntry.Open(), Encoding.UTF8))
                    {
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            if (i > 0)
                                writer.Write(',');
                            writer.Write(reader.GetName(i) );
                        }
                        writer.Write(Environment.NewLine);

                        while (reader.Read())
                        {
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                if (i > 0)
                                    writer.Write(',');
                                String v = reader[i].ToString();
                                if (v.Contains(',') || v.Contains('\n') || v.Contains('\r') || v.Contains('"'))
                                {
                                    writer.Write('"');
                                    writer.Write(v.Replace("\"", "\"\""));
                                    writer.Write('"');
                                }
                                else
                                {
                                    writer.Write(v);
                                }
                            }
                            writer.Write(Environment.NewLine);
                        }

                    }
                }


            }
            watch.Stop();
            return Convert.ToInt32(watch.ElapsedMilliseconds / 1000);

        }



    }
}
