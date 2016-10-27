using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Text;
enum ErrorCode : long { SUCCESS, FAIL };

namespace View_and_Save_ODBCData
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }

    public static class EventLog    // Write Test Log
    {
        public static string FilePath { get; set; }

        public static void AddLog(string format, params object[] arg)
        {
            AddLog(string.Format(format, arg));
        }

        public static void AddLog(string message)
        {
            if (string.IsNullOrEmpty(FilePath))
            {
                //FilePath = Directory.GetCurrentDirectory();
                FilePath = "C:\\WALogData\\";
            }
            string filename = FilePath +
                //string.Format("\\{0:yyyy}\\{0:MM}\\{0:yyyy-MM-dd}.txt", DateTime.Now);
            string.Format("{0:yyyy-MM-dd}.txt", DateTime.Now);
            FileInfo finfo = new FileInfo(filename);
            if (finfo.Directory.Exists == false)
            {
                finfo.Directory.Create();
            }
            string writeString = string.Format("{0:[yyyy/MM/dd HH:mm:ss]} {1}",
                DateTime.Now, message) + Environment.NewLine;
            File.AppendAllText(filename, writeString, Encoding.Unicode);
        }
    }

}


