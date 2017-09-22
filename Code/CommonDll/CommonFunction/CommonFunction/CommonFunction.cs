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
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System.Diagnostics;
using System.Collections.ObjectModel;   // for ReadOnlyCollection<T> use

namespace CommonFunction
{
    public class cWACommonFunction
    {
        cEventLog EventLog = new cEventLog();

        public bool Download(IWebDriver driver, string slanguage)
        {
            try
            {
                //Thread.Sleep(1000);
                //driver.SwitchTo().Window(driver.CurrentWindowHandle);
                //Thread.Sleep(1000);
                driver.SwitchTo().Frame("rightFrame");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("//tr[2]/td/a[3]/font")).Click();

                // get current window handle
                string originalHandle = driver.CurrentWindowHandle;

                // trigger popup window
                //driver.FindElement(By.Id("a_btn_viewMode")).Click();    // View dashboard
                //Thread.Sleep(3000);

                // get current all window handle
                string popupHandle = string.Empty;
                ReadOnlyCollection<string> windowHandles = driver.WindowHandles;

                foreach (string handle in windowHandles)
                {
                    if (handle != originalHandle)
                    {
                        popupHandle = handle; break;
                    }
                }

                //switch to new window 
                driver.SwitchTo().Window(popupHandle);
                Thread.Sleep(1000);
                driver.FindElement(By.Name("submit")).Click();

                //WebDriverWait _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                //System.Console.Write("a = {0} \n", _wait.Until(d => d.FindElement(By.CssSelector("font[class=e4]"))).Text);

                Boolean bDoneChk = true;
                ReadOnlyCollection<IWebElement> links = driver.FindElements(By.CssSelector("font[class=e3]"));
                for (int check = 0; check < 3; check++)
                {
                    if (links.Count < 5)
                    {
                        Thread.Sleep(5000);
                        links = driver.FindElements(By.CssSelector("font[class=e3]"));
                    }
                    else
                        break;
                }
                foreach (IWebElement link in links)
                {
                    //string text = link.Text;
                    //System.Console.Write("b = {0} \n", text);
                    
                    switch (slanguage)
                    {
                        case "ENG":
                            if ((link.Text == "Done") || (link.Text == "Primary SCADA Node restarted in Communication mode"))
                                bDoneChk = false;
                            break;
                        case "CHT":
                            if ((link.Text == "完成") || (link.Text == "再啟動主要監控節點為通信模式"))
                                bDoneChk = false;
                            break;
                        case "CHS":
                            if ((link.Text == "完成") || (link.Text == "再启动主要监控节点为通信模式"))
                                bDoneChk = false; 
                            break;
                        case "JPN":
                            if ((link.Text == "完了") || (link.Text == "ﾌﾟﾗｲﾏﾘSCADAﾉｰﾄﾞはｺﾐｭﾆｹｰｼｮﾝ ﾓｰﾄﾞで再起動しました。"))
                                bDoneChk = false;
                            break;
                        case "KRN":

                            break;
                        case "FRN":

                            break;

                        default:
                            if ((link.Text == "Done") || (link.Text == "Primary SCADA Node restarted in Communication mode"))
                                bDoneChk = false;
                            break;
                    }
                }

                string mqtt_msg = driver.FindElement(By.Id("mqtt_msg")).GetAttribute("style");
                if (mqtt_msg.IndexOf("none") < 0)
                {
                    EventLog.AddLog("[MQTT]Failed to Updated Config.");
                    return false;
                }

                string mqtt_msg_timeout = driver.FindElement(By.Id("mqtt_msg_timeout")).GetAttribute("style");
                if (mqtt_msg_timeout.IndexOf("none") < 0)
                {
                    EventLog.AddLog("[MQTT]Communication timeout.");
                    return false;
                }
                if (bDoneChk)
                {
                    EventLog.AddLog("Downland Fail.");
                    return false;
                }   
                
            }
            catch (Exception ex)
            {
                EventLog.AddLog(ex.ToString());
                return false;
            }
            return true;
        }

        

        public bool StartKernel(IWebDriver driver, string slanguage)
        {
            try
            {
                driver.SwitchTo().Frame("rightFrame");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("//tr[2]/td/a[5]/font")).Click();

                // get current window handle
                string originalHandle = driver.CurrentWindowHandle;

                // get current all window handle
                string popupHandle = string.Empty;
                ReadOnlyCollection<string> windowHandles = driver.WindowHandles;

                foreach (string handle in windowHandles)
                {
                    if (handle != originalHandle)
                    {
                        popupHandle = handle; break;
                    }
                }

                //switch to new window 
                driver.SwitchTo().Window(popupHandle);
                Thread.Sleep(1000);
                driver.FindElement(By.Name("submit")).Click();
                //Thread.Sleep(5000);
                Boolean bDoneChk = true;
                ReadOnlyCollection<IWebElement> links = driver.FindElements(By.CssSelector("font[class=e3]"));
                for (int check = 0; check < 3; check++)
                {
                    if (links.Count < 4)
                    {
                        Thread.Sleep(5000);
                        links = driver.FindElements(By.CssSelector("font[class=e3]"));
                    }
                    else
                        break;
                }
                foreach (IWebElement link in links)
                {
                    string text = link.Text;
                     System.Console.Write("b = {0} \n", text);


                    switch (slanguage)
                    {
                        case "ENG":
                            //if ((link.Text == "Done") || (link.Text == "Primary SCADA Node started in Communication mode") || (link.Text == "Primary SCADA Node is running. Nothing done."))
                            if ((link.Text == "Primary SCADA Node started in Communication mode") || (link.Text == "Primary SCADA Node is running. Nothing done."))
                                bDoneChk = false; EventLog.AddLog("Message: " + link.Text);
                            break;
                        case "CHT":
                            if ((link.Text == "啟動主要監控節點為通信模式") || (link.Text == "主要監控節點正在執行中,指令取消"))
                                bDoneChk = false; EventLog.AddLog("Message: " + link.Text);
                            break;
                        case "CHS":
                            if ((link.Text == "启动主要监控节点为通信模式") || (link.Text == "主要监控节点正在运行.指令取消."))
                                bDoneChk = false; EventLog.AddLog("Message: " + link.Text);
                            break;
                        case "JPN":
                            if ((link.Text == "ﾌﾟﾗｲﾏﾘSCADAﾉｰﾄﾞはｺﾐｭﾆｹｰｼｮﾝ ﾓｰﾄﾞで開始しました。") || (link.Text == "ﾌﾟﾗｲﾏﾘSCADAﾉｰﾄﾞは稼働しています。ｺﾏﾝﾄﾞを取り消します。"))
                                bDoneChk = false; EventLog.AddLog("Message: " + link.Text);
                            break;
                        case "KRN":
                            
                            break;
                        case "FRN":
                            
                            break;

                        default:
                            if ((link.Text == "Primary SCADA Node started in Communication mode") || (link.Text == "Primary SCADA Node is running. Nothing done."))
                                bDoneChk = false; EventLog.AddLog("Message: " + link.Text);
                            break;
                    }
                }
                if (bDoneChk)
                {
                    EventLog.AddLog("Start Kernel Fail.");
                    return false;
                }
                
            }
            catch (Exception ex)
            {
                EventLog.AddLog(ex.ToString());
                return false;
            }
            
            return true;
        }

        public bool StopKernel(IWebDriver driver, string slanguage)
        {
            try
            {
                driver.SwitchTo().Frame("rightFrame");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("//tr[2]/td/a[6]/font")).Click();

                // get current window handle
                string originalHandle = driver.CurrentWindowHandle;

                // get current all window handle
                string popupHandle = string.Empty;
                ReadOnlyCollection<string> windowHandles = driver.WindowHandles;

                foreach (string handle in windowHandles)
                {
                    if (handle != originalHandle)
                    {
                        popupHandle = handle; break;
                    }
                }

                //switch to new window 
                driver.SwitchTo().Window(popupHandle);
                Thread.Sleep(1000);
                driver.FindElement(By.Name("submit")).Click();

                Boolean bDoneChk = true;
                ReadOnlyCollection<IWebElement> links = driver.FindElements(By.CssSelector("font[class=e3]"));
                for (int check = 0; check < 3; check++)
                {
                    if (links.Count < 4)
                    {
                        Thread.Sleep(5000);
                        links = driver.FindElements(By.CssSelector("font[class=e3]"));
                    }
                    else
                        break;
                }
                foreach (IWebElement link in links)
                {
                    //string text = link.Text;
                    //System.Console.Write("b = {0} \n", text);
                    
                    switch (slanguage)
                    {
                        case "ENG":
                            if ((link.Text == "Primary SCADA Node stopped.") || (link.Text == "Primary SCADA Node is not running. Nothing done."))
                                bDoneChk = false; EventLog.AddLog("Message: " + link.Text);
                            break;
                        case "CHT":
                            if ((link.Text == "主要監控節點停止") || (link.Text == "主要監控節點沒有執行,指令取消"))
                                bDoneChk = false; EventLog.AddLog("Message: " + link.Text);
                            break;
                        case "CHS":
                            if ((link.Text == "主要监控节点停止") || (link.Text == "主要监控节点未运行.指令取消."))
                                bDoneChk = false; EventLog.AddLog("Message: " + link.Text);
                            break;
                        case "JPN":
                            if ((link.Text == "ﾌﾟﾗｲﾏﾘSCADAﾉｰﾄﾞ停止") || (link.Text == "ﾌﾟﾗｲﾏﾘSCADAﾉｰﾄﾞは稼働していません。ｺﾏﾝﾄﾞを取り消します。"))
                                bDoneChk = false; EventLog.AddLog("Message: " + link.Text);
                            break;
                        case "KRN":

                            break;
                        case "FRN":

                            break;

                        default:
                            if ((link.Text == "Primary SCADA Node stopped.") || (link.Text == "Primary SCADA Node is not running. Nothing done."))
                                bDoneChk = false; EventLog.AddLog("Message: " + link.Text);
                            break;
                    }
                }
                if (bDoneChk)
                {
                    EventLog.AddLog("Stop Kernel Fail.");
                    return false;
                }

            }
            catch (Exception ex)
            {
                EventLog.AddLog(ex.ToString());
                return false;
            }

            return true;
        }

        public void Download(IAdvSeleniumAPI api)
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

        public void StartKernel(IAdvSeleniumAPI api)
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

        public void StopKernel(IAdvSeleniumAPI api)
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

        public void StopKernel(IAdvSeleniumAPI api, bool bRedundancyTest)
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
