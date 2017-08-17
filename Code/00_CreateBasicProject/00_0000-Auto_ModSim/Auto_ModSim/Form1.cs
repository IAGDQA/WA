using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using ThirdPartyToolControl;
using iATester;
using System.Runtime.InteropServices;
using CommonFunction;
using System.Net.Sockets;
//step1. reference nmodbuspc.dll, and using the namespaces.
using Modbus.Device;      //for modbus master
using System.IO;

namespace Auto_ModSim
{
    public partial class Form1 : Form, iATester.iCom
    {
        [DllImport("User32.dll")]
        private static extern bool ShowWindowAsync(IntPtr hWnd, int cmdShow);
        private const int WS_SHOWNORMAL = 1;
        ModbusIpMaster master;
        TcpClient tcpClient;
        string ipAddress="127.0.0.1";
        int tcpPort = 502;

        cEventLog EventLog = new cEventLog();
        Stopwatch sw = new Stopwatch();
        bool bFinalResult = true;
        bool bPartResult = true;
        string sTestItemName = "Auto ModSim";

        private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
        private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
        internal const int Max_Rows_Val = 65535;

        //Send Log data to iAtester
        public event EventHandler<LogEventArgs> eLog = delegate { };
        //Send test result to iAtester
        public event EventHandler<ResultEventArgs> eResult = delegate { };
        //Send execution status to iAtester
        public event EventHandler<StatusEventArgs> eStatus = delegate { };

        public Form1()
        {
            InitializeComponent();
            Browser.SelectedIndex = 0;
            Text = string.Format("Advantech WebAccess Auto Test ( {0} )", sTestItemName);
            try
            {
                m_DataGridViewCtrlAddDataRow = new DataGridViewCtrlAddDataRow(DataGridViewCtrlAddNewRow);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            //if (System.IO.File.Exists(sIniFilePath))
            //{
            //    EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
            //    InitialRequiredInfo(sIniFilePath);
            //}
        }

        public void StartTest()
        {
            long lErrorCode = 0;
            EventLog.AddLog(string.Format("==={0} test start (by iATester)===", sTestItemName));
            //if (System.IO.File.Exists(sIniFilePath))    // 再load一次
            //{
            //    EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
            //    InitialRequiredInfo(sIniFilePath);
            //}
            //EventLog.AddLog("Project= " + ProjectName.Text);
            //EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load();
            EventLog.AddLog(string.Format("==={0} test end (by iATester)===", sTestItemName));

            if (lErrorCode == 0)
                eResult(this, new ResultEventArgs(iResult.Pass));
            else
                eResult(this, new ResultEventArgs(iResult.Fail));

            eStatus(this, new StatusEventArgs(iStatus.Completion));
        }


        private void Form1_Load(object sender, EventArgs e)
        {


            //int aa = tpc.F_GetWindowText(iEnterText2,"",0);
            //tpc.F_PostMessage(iEnterText2, tpc.V_WM_KEYDOWN, tpc.V_VK_RETURN, 0);
            //else
            //EventLog.AddLog("Cannot get Recipe List handle");

            //Process[] processes = Process.GetProcessesByName("ModSim32");
            ////IntPtr player;
            //foreach (Process p in processes)
            //{
            //    // 關閉目前程序前先等待 1000 毫秒
            //    //p.WaitForExit(1000);
            //    //p.CloseMainWindow();
            //    IntPtr player = p.Handle;
            //    SetForegroundWindow(player);
            //}

            //Process ModSim = new Process();
            //// FileName 是要執行的檔案
            //ModSim.StartInfo.FileName = @"C:\ModSim\ModSim32.exe";
            //ModSim.Start();

            //cThirdPartyToolControl tpc = new cThirdPartyToolControl();
            //int NotepadHwnd = FindWindow("ModSim32", null);
            //int iEnterText = tpc.F_FindWindowEx(iLoginKeyboard_Handle, 0, "Edit", "");
            //SetForegroundWindow(player);
            //IntPtr player = tpc.F_FindWindow("CalcFrame", null);
            //int gm = GetMenu(player);
            //gm = GetSubMenu(gm, 0);
            //int id = GetMenuItemID(gm, 1);
            //System.Threading.Thread.Sleep(1000);
            //PostMessage(player, tpc.V_VK_F10, id, 0);
            //System.Threading.Thread.Sleep(3000);
            //string aaaa ="";
            //PostMessage(NotepadHwnd, tpc.V_VK_F10, id, 0);
            //System.Threading.Thread.Sleep(1000);
            //tpc.F_KeybdEvent(tpc.V_VK_CONTROL, 0, tpc.V_KEYEVENTF_EXTENDEDKEY, 0);
            //tpc.F_PostMessage(iLoginKeyboard_Handle, tpc.V_WM_KEYDOWN, tpc.V_VK_F10, 0);
            //System.Threading.Thread.Sleep(1000);
            //tpc.F_KeybdEvent(tpc.V_VK_CONTROL, 0, tpc.V_KEYEVENTF_KEYUP, 0);
            //System.Threading.Thread.Sleep(1000);

            //tpc.F_KeybdEvent(tpc.V_VK_CONTROL, 0, tpc.V_KEYEVENTF_KEYUP, 0);
            //System.Threading.Thread.Sleep(1000);
            //tpc.F_KeybdEvent(0x004F, 0, tpc.V_KEYEVENTF_EXTENDEDKEY, 0);
            //System.Threading.Thread.Sleep(1000);
        }

        private void Start_Click(object sender, EventArgs e)
        {
            long lErrorCode = 0;
            EventLog.AddLog(string.Format("==={0} test start===", sTestItemName));
            //CheckifIniFileChange();
            //EventLog.AddLog("Project= " + ProjectName.Text);
            //EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode= Form1_Load();
            PrintStep("Auto ModSim", "open, load, connection and verification", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());
            EventLog.AddLog(string.Format("==={0} test end===", sTestItemName));
            //EventLog.AddLog(string.Format("==={0} test start (by iATester)===", sTestItemName));
            //cThirdPartyToolControl tpc = new cThirdPartyToolControl();

            //Process[] processes = Process.GetProcessesByName("ModSim32");
            ////IntPtr player;
            //foreach (Process p in processes)
            //{
            //    // 關閉目前程序前先等待 1000 毫秒
            //    p.WaitForExit(1000);
            //    p.CloseMainWindow();
            //}
            //EventLog.AddLog("Check existed ModSim and Close it");

            //Process ModSim = new Process();
            //// FileName 是要執行的檔案
            //ModSim.StartInfo.FileName = @"C:\ModSim\ModSim32.exe";
            //ModSim.Start();

            //EventLog.AddLog("Excute ModSim");

            ////EventLog.AddLog("wait 2 sec");
            //System.Threading.Thread.Sleep(3000);
            ////EventLog.AddLog("enter");
            //tpc.F_KeybdEvent(tpc.V_VK_ESCAPE, 0, tpc.V_KEYEVENTF_EXTENDEDKEY, 0);
            ////EventLog.AddLog("wait 1 sec");
            //System.Threading.Thread.Sleep(1000);

            //foreach (Process pTarget in Process.GetProcesses())
            //{
            //    if (pTarget.ProcessName == "ModSim32")  // 取得處理序名稱並與指定程序名稱比較
            //    {
            //        //HandleRunningInstance(pTarget);
            //        ShowWindowAsync(pTarget.MainWindowHandle, WS_SHOWNORMAL);
            //        tpc.F_SetForegroundWindow((int)pTarget.MainWindowHandle);
            //    }
            //}

            //EventLog.AddLog("Load script");

            //string[] file_name = { "CoilStatus_250_Sync_20160331", "HoldingRegister_250_Sync_20160401", "InputRegister_250_Sync_20160331", "InputStatus_250_Sync_20160401" };
            ////string[] file_name = { "CoilStatus_250_Sync_20160331" };

            //for (int i = 0; i < file_name.Length; i++)
            //{

            //    System.Threading.Thread.Sleep(2000);
            //    SendKeys.SendWait("^o");
            //    System.Threading.Thread.Sleep(1000);

            //    int iRecipeList_Handle = tpc.F_FindWindow("#32770", "Open");
            //    StringBuilder lpString = new StringBuilder(10);
            //    int bb = tpc.F_GetWindowText(iRecipeList_Handle, lpString, 100);

            //    //int iEnterText1 = tpc.F_FindWindowEx(iRecipeList_Handle, 0, "ComboBox", "");
            //    //StringBuilder test = new StringBuilder(10);
            //    //System.Threading.Thread.Sleep(1000);
            //    //tpc.F_PostMessage(iEnterText1, tpc.V_WM_CHAR, 'a', 0);
            //    //System.Threading.Thread.Sleep(1000);
            //    //int aa = tpc.F_GetWindowText(iEnterText1, test, 0);
            //    //string aaa = test.ToString();

            //    int iEnterText2 = tpc.F_FindWindowEx(iRecipeList_Handle, 0, "Edit", "");
            //    if (iEnterText2 > 0)
            //    {

            //        byte[] ch = (ASCIIEncoding.ASCII.GetBytes(file_name[i]));
            //        for (int j = 0; j < ch.Length; j++)
            //        {
            //            //SendMessage(PW, WM_CHAR, ch, 0);
            //            tpc.F_PostMessage(iEnterText2, tpc.V_WM_CHAR, ch[j], 0);
            //            //System.Threading.Thread.Sleep(100);
            //        }
            //        tpc.F_PostMessage(iEnterText2, tpc.V_WM_KEYDOWN, tpc.V_VK_RETURN, 0);

            //        //tpc.F_PostMessage(iEnterText2, tpc.V_WM_CHAR, 'a', 0);

            //    }
            //}

            //EventLog.AddLog("connect to TCP");

            //System.Threading.Thread.Sleep(2000);
            //SendKeys.SendWait("%c");
            //System.Threading.Thread.Sleep(1000);
            //SendKeys.SendWait("%c");
            //System.Threading.Thread.Sleep(1000);
            //tpc.F_KeybdEvent(tpc.V_VK_RETURN, 0, tpc.V_KEYEVENTF_EXTENDEDKEY, 0);
            //System.Threading.Thread.Sleep(1000);
            //tpc.F_KeybdEvent(tpc.V_VK_UP, 0, tpc.V_KEYEVENTF_EXTENDEDKEY, 0);
            //System.Threading.Thread.Sleep(1000);
            //tpc.F_KeybdEvent(tpc.V_VK_RETURN, 0, tpc.V_KEYEVENTF_EXTENDEDKEY, 0);
            //System.Threading.Thread.Sleep(2000);
            //tpc.F_KeybdEvent(tpc.V_VK_RETURN, 0, tpc.V_KEYEVENTF_EXTENDEDKEY, 0);
        }
        //public static void HandleRunningInstance(Process instance)
        //{
        //    // 相同時透過ShowWindowAsync還原，以及SetForegroundWindow將程式至於前景
        //    ShowWindowAsync(instance.MainWindowHandle, WS_SHOWNORMAL);
        //    tcp.F_SetForegroundWindow(instance.MainWindowHandle);
        //}
        long Form1_Load()
        {
            cThirdPartyToolControl tpc = new cThirdPartyToolControl();
            try
            {
                Process[] processes = Process.GetProcessesByName("ModSim32");
                foreach (Process p in processes)
                {
                    // 關閉目前程序前先等待 1000 毫秒
                    p.WaitForExit(1000);
                    p.CloseMainWindow();
                }
                EventLog.AddLog("Check existed ModSim and Close it");
                System.Threading.Thread.Sleep(1000);

                string str = this.GetType().Assembly.Location;
                str= str.Substring(0, str.LastIndexOf(@"\"));
                string sourceDirName = @"\ModSim";
                string destDirName=@"c:\ModSim";
                bool copySubDirs = true;
                DirectoryCopy(str + sourceDirName, destDirName, copySubDirs);
            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred Delete ModSim and Move it to C:\ : " + ex.ToString());
                Result.Text = "FAIL!!";
                Result.ForeColor = Color.Red;
                EventLog.AddLog("Test Result: FAIL!!");
                return -1;
            }

            try
            {
                Process ModSim = new Process();
                // FileName 是要執行的檔案
                ModSim.StartInfo.FileName = @"C:\ModSim\ModSim32.exe";
                ModSim.Start();

                EventLog.AddLog("Excute ModSim");
            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred Excute ModSim: " + ex.ToString());
                Result.Text = "FAIL!!";
                Result.ForeColor = Color.Red;
                EventLog.AddLog("Test Result: FAIL!!");
                return -1;
            }

            try
            {
                System.Threading.Thread.Sleep(3000);
                //EventLog.AddLog("enter");
                tpc.F_KeybdEvent(tpc.V_VK_ESCAPE, 0, tpc.V_KEYEVENTF_EXTENDEDKEY, 0);
                //EventLog.AddLog("wait 1 sec");
                System.Threading.Thread.Sleep(1000);

                foreach (Process pTarget in Process.GetProcesses())
                {
                    if (pTarget.ProcessName == "ModSim32")  // 取得處理序名稱並與指定程序名稱比較
                    {
                        //HandleRunningInstance(pTarget);
                        ShowWindowAsync(pTarget.MainWindowHandle, WS_SHOWNORMAL);
                        tpc.F_SetForegroundWindow((int)pTarget.MainWindowHandle);
                    }
                }

                EventLog.AddLog("Load script");

                string[] file_name = { "CoilStatus_250_Sync_20160331", "HoldingRegister_250_Sync_20160401", "InputRegister_250_Sync_20160331", "InputStatus_250_Sync_20160401" };
                //string[] file_name = { "CoilStatus_250_Sync_20160331" };

                for (int i = 0; i < file_name.Length; i++)
                {

                    System.Threading.Thread.Sleep(2000);
                    SendKeys.SendWait("^o");
                    System.Threading.Thread.Sleep(1000);

                    int iRecipeList_Handle = tpc.F_FindWindow("#32770", "Open");
                    StringBuilder lpString = new StringBuilder(10);
                    int bb = tpc.F_GetWindowText(iRecipeList_Handle, lpString, 100);
                    int iEnterText2 = tpc.F_FindWindowEx(iRecipeList_Handle, 0, "Edit", "");
                    if (iEnterText2 > 0)
                    {
                        byte[] ch = (ASCIIEncoding.ASCII.GetBytes(file_name[i]));
                        for (int j = 0; j < ch.Length; j++)
                        {
                            //SendMessage(PW, WM_CHAR, ch, 0);
                            tpc.F_PostMessage(iEnterText2, tpc.V_WM_CHAR, ch[j], 0);
                            //System.Threading.Thread.Sleep(100);
                        }
                        tpc.F_PostMessage(iEnterText2, tpc.V_WM_KEYDOWN, tpc.V_VK_RETURN, 0);

                        //tpc.F_PostMessage(iEnterText2, tpc.V_WM_CHAR, 'a', 0);
                    }
                }
            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred Load file: " + ex.ToString());
                Result.Text = "FAIL!!";
                Result.ForeColor = Color.Red;
                EventLog.AddLog("Test Result: FAIL!!");
                return -1;
            }

            try
            {
                EventLog.AddLog("connect to TCP");

                System.Threading.Thread.Sleep(2000);
                SendKeys.SendWait("%c");
                System.Threading.Thread.Sleep(1000);
                SendKeys.SendWait("%c");
                System.Threading.Thread.Sleep(1000);
                tpc.F_KeybdEvent(tpc.V_VK_RETURN, 0, tpc.V_KEYEVENTF_EXTENDEDKEY, 0);
                System.Threading.Thread.Sleep(1000);
                tpc.F_KeybdEvent(tpc.V_VK_UP, 0, tpc.V_KEYEVENTF_EXTENDEDKEY, 0);
                System.Threading.Thread.Sleep(1000);
                tpc.F_KeybdEvent(tpc.V_VK_RETURN, 0, tpc.V_KEYEVENTF_EXTENDEDKEY, 0);
                System.Threading.Thread.Sleep(2000);
                tpc.F_KeybdEvent(tpc.V_VK_RETURN, 0, tpc.V_KEYEVENTF_EXTENDEDKEY, 0);
            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred connect to TCP: " + ex.ToString());
                Result.Text = "FAIL!!";
                Result.ForeColor = Color.Red;
                EventLog.AddLog("Test Result: FAIL!!");
                return -1;
            }

            try
            {
                if (master != null)
                    master.Dispose();
                if (tcpClient != null)
                    tcpClient.Close();
                tcpClient = new TcpClient();
                IAsyncResult asyncResult = tcpClient.BeginConnect(ipAddress, tcpPort, null, null);
                asyncResult.AsyncWaitHandle.WaitOne(3000, true); //wait for 3 sec
                if (!asyncResult.IsCompleted)
                {
                    tcpClient.Close();
                    EventLog.AddLog("Cannot connect to ModSim.");
                    return -1;
                }
                //tcpClient = new TcpClient(ipAddress, tcpPort);

                // create Modbus TCP Master by the tcp client
                //document->Modbus.Device.Namespace->ModbusIpMaster Class->Create Method
                master = ModbusIpMaster.CreateIp(tcpClient);
                master.Transport.Retries = 0;   //don't have to do retries
                master.Transport.ReadTimeout = 1500;
                //this.Text = "On line " + DateTime.Now.ToString();
                EventLog.AddLog("Connect to ModSim.");

                //read DI(1xxxx), start address=0, points=4
                byte slaveID = 1;
                bool[] status = master.ReadInputs(slaveID, 0, 1);
                if ((status[0] == true) || (status[0] == false))
                {

                }
                else
                {
                    EventLog.AddLog("read DI(10001) fail");
                    return -1;
                }
                //read DO(00001), start address=0, points=1
                bool[] coils = master.ReadCoils(slaveID, 0, 1);
                if ((coils[0] == true) || (coils[0] == false))
                {

                }
                else
                {
                    EventLog.AddLog("read DO(00001) fail");
                    return -1;
                }
                //read AI(30001), start address=0, points=1
                ushort[] register = master.ReadInputRegisters(1, 0, 1);
                if (register[0] > 0)
                {

                }
                else
                {
                    EventLog.AddLog("read AI(30001) fail");
                    return -1;
                }
                //read AO(40001), start address=0, points=1
                ushort[] holding_register = master.ReadHoldingRegisters(1, 0, 1);
                if (holding_register[0] > 0)
                {

                }
                else
                {
                    EventLog.AddLog("read AO(40001) fail");
                    return -1;
                }
            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred verification: " + ex.ToString());
                Result.Text = "FAIL!!";
                Result.ForeColor = Color.Red;
                EventLog.AddLog("Test Result: FAIL!!");
                return -1;
            }

            #region Result judgement
            int iTotalSeleniumAction = dataGridView1.Rows.Count;
            for (int i = 0; i < iTotalSeleniumAction - 1; i++)
            {
                DataGridViewRow row = dataGridView1.Rows[i];
                string sSeleniumResult = row.Cells[2].Value.ToString();
                if (sSeleniumResult != "PASS")
                {
                    bFinalResult = false;
                    EventLog.AddLog("Test Fail !!");
                    EventLog.AddLog("Fail TestItem = " + row.Cells[0].Value.ToString());
                    EventLog.AddLog("BrowserAction = " + row.Cells[1].Value.ToString());
                    EventLog.AddLog("Result = " + row.Cells[2].Value.ToString());
                    EventLog.AddLog("ErrorCode = " + row.Cells[3].Value.ToString());
                    EventLog.AddLog("ExeTime(ms) = " + row.Cells[4].Value.ToString());
                    break;
                }
            }

            if (bFinalResult)
            {
                Result.Text = "PASS!!";
                Result.ForeColor = Color.Green;
                EventLog.AddLog("Test Result: PASS!!");
                return 0;
            }
            else
            {
                Result.Text = "FAIL!!";
                Result.ForeColor = Color.Red;
                EventLog.AddLog("Test Result: FAIL!!");
                return -1;
            }
            #endregion
        }
        private void PrintStep(string sTestItem, string sDescription, bool bResult, string sErrorCode, string sExTime)
        {
            EventLog.AddLog(string.Format("UI Result: {0},{1},{2},{3},{4}", sTestItem, sDescription, bResult, sErrorCode, sExTime));

            DataGridViewRow dgvRow;
            DataGridViewCell dgvCell;

            dgvRow = new DataGridViewRow();

            if (bResult == false)
                dgvRow.DefaultCellStyle.ForeColor = Color.Red;

            dgvCell = new DataGridViewTextBoxCell(); //Column Time

            dgvCell.Value = sTestItem;
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            dgvCell.Value = sDescription;
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            if (bResult)
                dgvCell.Value = "PASS";
            else
                dgvCell.Value = "FAIL";
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            dgvCell.Value = sErrorCode;
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            dgvCell.Value = sExTime;
            dgvRow.Cells.Add(dgvCell);

            m_DataGridViewCtrlAddDataRow(dgvRow);
        }
        private void DataGridViewCtrlAddNewRow(DataGridViewRow i_Row)
        {
            if (this.dataGridView1.InvokeRequired)
            {
                this.dataGridView1.Invoke(new DataGridViewCtrlAddDataRow(DataGridViewCtrlAddNewRow), new object[] { i_Row });
                return;
            }

            this.dataGridView1.Rows.Insert(0, i_Row);
            if (dataGridView1.Rows.Count > Max_Rows_Val)
            {
                dataGridView1.Rows.RemoveAt((dataGridView1.Rows.Count - 1));
            }
            this.dataGridView1.Update();
        }

        private static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            // Get the subdirectories for the specified directory.
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException(
                    "Source directory does not exist or could not be found: "
                    + sourceDirName);
            }

            DirectoryInfo[] dirs = dir.GetDirectories();
            // If the destination directory doesn't exist, create it.
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
            }

            // Get the files in the directory and copy them to the new location.
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, true);
            }

            // If copying subdirectories, copy them and their contents to new location.
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }
    }
}
