using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;                // for AddLog
using System.Drawing;           // for PrintScreen
using System.Windows.Forms;     // for PrintScreen
using AdvWebUIAPI;

namespace CommonFunction
{
    public class cWACommonFunction
    {
        cEventLog EventLog = new cEventLog();

        public void StartDownload(IAdvSeleniumAPI api)
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//tr[2]/td/a[3]/font").Click();    // "Download" click
            Thread.Sleep(2000);
            EventLog.AddLog("Find pop up download window handle");
            string main; object subobj;                     // Find pop up download window handle
            api.GetWinHandle(out main, out subobj);
            IEnumerator<String> windowIterator = (IEnumerator<String>)subobj;

            List<string> items = new List<string>();
            while (windowIterator.MoveNext())
                items.Add(windowIterator.Current);

            EventLog.AddLog("Main window handle= " + main);
            EventLog.AddLog("Window handle list items[0]= " + items[0]);
            EventLog.AddLog("Window handle list items[1]= " + items[1]);
            if (main != items[1])
            {
                EventLog.AddLog("Switch to items[1]");
                api.SwitchToWinHandle(items[1]);
            }
            else
            {
                EventLog.AddLog("Switch to items[0]");
                api.SwitchToWinHandle(items[0]);
            }
            api.ByName("submit").Enter("").Submit().Exe();

            EventLog.AddLog("Start to download and wait 80 seconds...");
            Thread.Sleep(80000);    // Wait 80s for Download finish
            EventLog.AddLog("It's been wait 80 seconds");
            EventLog.PrintScreen("Download result");
            api.Close();
            EventLog.AddLog("Close download window and switch to main window");
            api.SwitchToWinHandle(main);
        }

        private void StartKernel(IAdvSeleniumAPI api)
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//tr[2]/td/a[5]/font").Click();    // start kernel
            Thread.Sleep(2000);
            EventLog.AddLog("Find pop up StartNode window handle");
            string main; object subobj;
            api.GetWinHandle(out main, out subobj);
            IEnumerator<String> windowIterator = (IEnumerator<String>)subobj;

            List<string> items = new List<string>();
            while (windowIterator.MoveNext())
                items.Add(windowIterator.Current);

            EventLog.AddLog("Main window handle= " + main);
            EventLog.AddLog("Window handle list items[0]= " + items[0]);
            EventLog.AddLog("Window handle list items[1]= " + items[1]);
            if (main != items[1])
            {
                EventLog.AddLog("Switch to items[1]");
                api.SwitchToWinHandle(items[1]);
            }
            else
            {
                EventLog.AddLog("Switch to items[0]");
                api.SwitchToWinHandle(items[0]);
            }
            api.ByName("submit").Enter("").Submit().Exe();

            EventLog.AddLog("Start node and wait 30 seconds...");
            Thread.Sleep(30000);    // Wait 30s for start kernel finish
            EventLog.AddLog("It's been wait 30 seconds");
            EventLog.PrintScreen("Start Node result");
            api.Close();
            EventLog.AddLog("Close start node window and switch to main window");
            api.SwitchToWinHandle(main);        // switch back to original window
        }

        private void StopKernel(IAdvSeleniumAPI api)
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//tr[2]/td/a[6]/font").Click();    // Stop kernel
            Thread.Sleep(2000);

            EventLog.AddLog("Find pop up StopNode window handle");
            string main; object subobj;
            api.GetWinHandle(out main, out subobj);
            IEnumerator<String> windowIterator = (IEnumerator<String>)subobj;

            List<string> items = new List<string>();
            while (windowIterator.MoveNext())
                items.Add(windowIterator.Current);

            EventLog.AddLog("Main window handle= " + main);
            EventLog.AddLog("Window handle list items[0]= " + items[0]);
            EventLog.AddLog("Window handle list items[1]= " + items[1]);
            if (main != items[1])
            {
                EventLog.AddLog("Switch to items[1]");
                api.SwitchToWinHandle(items[1]);
            }
            else
            {
                EventLog.AddLog("Switch to items[0]");
                api.SwitchToWinHandle(items[0]);
            }
            /*
            if (bRedundancyTest == true)
            {
                Thread.Sleep(500);
                api.ByXpath("(//input[@name='SECONDARY_CONTROL'])[2]").Click();
                Thread.Sleep(1000);
            }
            api.ByName("submit").Enter("").Submit().Exe();

            if (bRedundancyTest == true)
            {
                EventLog.AddLog("Stop node and wait 100 seconds for redundancy test...");
                Thread.Sleep(100000);    // Wait 100s for Stop kernel finish
                EventLog.AddLog("It's been wait 100 seconds");
            }
            else
            {
                EventLog.AddLog("Stop node and wait 30 seconds...");
                Thread.Sleep(30000);    // Wait 30s for Stop kernel finish
                EventLog.AddLog("It's been wait 30 seconds");
            }
            */
            EventLog.PrintScreen("Stop Node result");
            api.Close();
            EventLog.AddLog("Close stop node window and switch to main window");
            api.SwitchToWinHandle(main);        // switch back to original window
        }

        private void StopKernel(IAdvSeleniumAPI api, bool bRedundancyTest)
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//tr[2]/td/a[6]/font").Click();    // Stop kernel
            Thread.Sleep(2000);

            EventLog.AddLog("Find pop up StopNode window handle");
            string main; object subobj;
            api.GetWinHandle(out main, out subobj);
            IEnumerator<String> windowIterator = (IEnumerator<String>)subobj;

            List<string> items = new List<string>();
            while (windowIterator.MoveNext())
                items.Add(windowIterator.Current);

            EventLog.AddLog("Main window handle= " + main);
            EventLog.AddLog("Window handle list items[0]= " + items[0]);
            EventLog.AddLog("Window handle list items[1]= " + items[1]);
            if (main != items[1])
            {
                EventLog.AddLog("Switch to items[1]");
                api.SwitchToWinHandle(items[1]);
            }
            else
            {
                EventLog.AddLog("Switch to items[0]");
                api.SwitchToWinHandle(items[0]);
            }
            
            if (bRedundancyTest == true)
            {
                Thread.Sleep(500);
                api.ByXpath("(//input[@name='SECONDARY_CONTROL'])[2]").Click();
                Thread.Sleep(1000);
            }
            api.ByName("submit").Enter("").Submit().Exe();

            if (bRedundancyTest == true)
            {
                EventLog.AddLog("Stop node and wait 100 seconds for redundancy test...");
                Thread.Sleep(100000);    // Wait 100s for Stop kernel finish
                EventLog.AddLog("It's been wait 100 seconds");
            }
            else
            {
                EventLog.AddLog("Stop node and wait 30 seconds...");
                Thread.Sleep(30000);    // Wait 30s for Stop kernel finish
                EventLog.AddLog("It's been wait 30 seconds");
            }
            
            EventLog.PrintScreen("Stop Node result");
            api.Close();
            EventLog.AddLog("Close stop node window and switch to main window");
            api.SwitchToWinHandle(main);        // switch back to original window
        }   // for Redundancy test

    }

    public class cEventLog    // Write Test Log
    {
        public string FilePath { get; set; }

        public void AddLog(string format, params object[] arg)
        {
            AddLog(string.Format(format, arg));
        }

        public void AddLog(string message)
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

        public void PrintScreen(string sFileName)
        {
            if (string.IsNullOrEmpty(FilePath))
            {
                //FilePath = Directory.GetCurrentDirectory();
                FilePath = "C:\\WALogData\\";
            }
            Bitmap myImage = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics g = Graphics.FromImage(myImage);
            g.CopyFromScreen(new Point(0, 0), new Point(0, 0), new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height));
            IntPtr dc1 = g.GetHdc();
            g.ReleaseHdc(dc1);
            //myImage.Save(@"c:\screen0.jpg");
            myImage.Save(string.Format("{0}\\{1}_{2:yyyyMMdd_hhmmss}.jpg", FilePath, sFileName, DateTime.Now));
        }
    }
}
