using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
namespace DataPie
{
    static class Program
    {

        //public static login log = null; 
  
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            int ProceedingCount = 0;
            Process[] ProceddingCon = Process.GetProcesses();
            foreach (Process IsProcedding in ProceddingCon)
            {
                if (IsProcedding.ProcessName == Process.GetCurrentProcess().ProcessName)
                {
                    ProceedingCount += 1;
                }
            }
            if (ProceedingCount > 1)
            {
                MessageBox.Show("该系统已经在运行中。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {

                Application.EnableVisualStyles();
                //Application.SetCompatibleTextRenderingDefault(false);
                //log = new login();
                Application.Run(new login());
            }
        }
    }
}
