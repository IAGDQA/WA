using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using AdvWebUIAPI;
using System.Runtime.InteropServices;
using System.Diagnostics;
using ThirdPartyToolControl;
using iATester;

namespace PlugandPlay_TagInfoSyncTest
{
    public partial class Form1 : Form, iATester.iCom
    {
        IAdvSeleniumAPI api;
        IAdvSeleniumAPI api2;
        cThirdPartyToolControl tpc = new cThirdPartyToolControl();

        private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
        private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
        internal const int Max_Rows_Val = 65535;
        string baseUrl, baseUrl2;
        string sIniFilePath = @"C:\WebAccessAutoTestSetting.ini";
        string slanguage;

        //Send Log data to iAtester
        public event EventHandler<LogEventArgs> eLog = delegate { };
        //Send test result to iAtester
        public event EventHandler<ResultEventArgs> eResult = delegate { };
        //Send execution status to iAtester
        public event EventHandler<StatusEventArgs> eStatus = delegate { };

        public void StartTest()
        {
            //Add test code
            long lErrorCode = (long)ErrorCode.SUCCESS;
            EventLog.AddLog("===PlugandPlay_TagInfoSyncTest start (by iATester)===");
            if (System.IO.File.Exists(sIniFilePath))    // 再load一次
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, ProjectName2.Text, WebAccessIP.Text, WebAccessIP2.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===PlugandPlay_TagInfoSyncTest end (by iATester)===");

            if (lErrorCode == 0)
                eResult(this, new ResultEventArgs(iResult.Pass));
            else
                eResult(this, new ResultEventArgs(iResult.Fail));

            eStatus(this, new StatusEventArgs(iStatus.Completion));
        }

        public Form1()
        {
            InitializeComponent();
            try
            {
                m_DataGridViewCtrlAddDataRow = new DataGridViewCtrlAddDataRow(DataGridViewCtrlAddNewRow);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            Browser.SelectedIndex = 0;
            if (System.IO.File.Exists(sIniFilePath))
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
        }

        long Form1_Load(string sProjectName, string sProjectName2, string sWebAccessIP, string sWebAccessIP2, string sTestLogFolder, string sBrowser)
        {
            baseUrl = "http://" + sWebAccessIP;
            baseUrl2 = "http://" + sWebAccessIP2;

            // Step1: Start the kernel of ground PC
            GroundPCStartKernel(sBrowser, sProjectName, sWebAccessIP, sTestLogFolder);

            // Step2: Start the kernel of cloud PC and view and save tag info
            ViewandSaveCloudTagInfo(sBrowser, sProjectName2, sWebAccessIP2, sTestLogFolder);

            bool bSeleniumResult = true;
            int iTotalSeleniumAction = dataGridView1.Rows.Count;
            for (int i = 0; i < iTotalSeleniumAction - 1; i++)
            {
                DataGridViewRow row = dataGridView1.Rows[i];
                string sSeleniumResult = row.Cells[2].Value.ToString();
                if (sSeleniumResult != "pass")
                {
                    bSeleniumResult = false;
                    EventLog.AddLog("Test Fail !!");
                    EventLog.AddLog("Fail TestItem = " + row.Cells[0].Value.ToString());
                    EventLog.AddLog("BrowserAction = " + row.Cells[1].Value.ToString());
                    EventLog.AddLog("Result = " + row.Cells[2].Value.ToString());
                    EventLog.AddLog("ErrorCode = " + row.Cells[3].Value.ToString());
                    EventLog.AddLog("ExeTime(ms) = " + row.Cells[4].Value.ToString());
                    break;
                }
            }

            if (bSeleniumResult)
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

            //return 0;
        }

        private void GroundPCStartKernel(string sBrowser, string sProjectName, string sWebAccessIP, string sTestLogFolder)
        {
            if (sBrowser == "Internet Explorer")
            {
                EventLog.AddLog("<GroundPC> Browser= Internet Explorer");
                api = new AdvSeleniumAPI("IE", "");
                System.Threading.Thread.Sleep(1000);
            }
            else if (sBrowser == "Mozilla FireFox")
            {
                EventLog.AddLog("<GroundPC> Browser= Mozilla FireFox");
                api = new AdvSeleniumAPI("FireFox", "");
                System.Threading.Thread.Sleep(1000);
            }

            api.LinkWebUI(baseUrl + "/broadWeb/bwconfig.asp?username=admin");
            api.ById("userField").Enter("").Submit().Exe();
            PrintStep(api, "<GroundPC> Login WebAccess");

            // Configure project by project name
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
            PrintStep(api, "<GroundPC> Configure project");

            EventLog.AddLog("<GroundPC> Start Kernel");
            StartNode(api, sTestLogFolder);

            api.Quit();
            PrintStep(api, "<GroundPC> Quit browser");
        }

        private void ViewandSaveCloudTagInfo(string sBrowser, string sProjectName, string sWebAccessIP, string sTestLogFolder)
        {
            if (sBrowser == "Internet Explorer")
            {
                EventLog.AddLog("<CloudPC> Browser= Internet Explorer");
                api2 = new AdvSeleniumAPI("IE", "");
                System.Threading.Thread.Sleep(1000);
            }
            else if (sBrowser == "Mozilla FireFox")
            {
                EventLog.AddLog("<CloudPC> Browser= Mozilla FireFox");
                api2 = new AdvSeleniumAPI("FireFox", "");
                System.Threading.Thread.Sleep(1000);
            }
            EventLog.AddLog("<CloudPC> Capture the project manager page");
            api2.LinkWebUI(baseUrl2 + "/broadWeb/bwconfig.asp?username=admin");
            api2.ById("userField").Enter("").Submit().Exe();
            PrintStep(api2, "<CloudPC> Login WebAccess");
            

            // Configure project by project name
            EventLog.AddLog("<CloudPC> Capture the configure project page");
            api2.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
            PrintStep(api2, "<CloudPC> Configure project");

            // Start kernel
            EventLog.AddLog("<CloudPC> Start kernel");
            StartNode(api2, sTestLogFolder);

            Thread.Sleep(5000);

            // Start view
            EventLog.AddLog("<CloudPC> start view..");
            api2.SwitchToCurWindow(0);
            api2.SwitchToFrame("rightFrame", 0);
            api2.ByXpath("//tr[2]/td/a/font").Click();
            PrintStep(api2, "Start View");

            // Control browser
            int iIE_Handl, iIE_Handl_2, iIE_Handl_3, iIE_Handl_4, iIE_Handl_5, iIE_Handl_6, iIE_Handl_7, iWA_MainPage = 0;
            switch (slanguage)
            {
                case "ENG":
                    iIE_Handl = tpc.F_FindWindow("IEFrame", "Node : CTestSCADA - main:untitled");
                    iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
                    iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "Node : CTestSCADA - Internet Explorer");
                    iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
                    iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");
                    iIE_Handl_6 = tpc.F_FindWindowEx(iIE_Handl_5, 0, "AfxOleControl42s", "");
                    iIE_Handl_7 = tpc.F_FindWindowEx(iIE_Handl_6, 0, "AfxWnd42s", "");
                    iWA_MainPage = tpc.F_FindWindowEx(iIE_Handl_7, 0, "ActXBroadWinBwviewWClass", "Advantech View 001 - main:untitled");
                    break;
                case "CHT":
                    iIE_Handl = tpc.F_FindWindow("IEFrame", "節點 : CTestSCADA - main:untitled");
                    iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
                    iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "節點 : CTestSCADA - Internet Explorer");
                    iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
                    iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");
                    iIE_Handl_6 = tpc.F_FindWindowEx(iIE_Handl_5, 0, "AfxOleControl42s", "");
                    iIE_Handl_7 = tpc.F_FindWindowEx(iIE_Handl_6, 0, "AfxWnd42s", "");
                    iWA_MainPage = tpc.F_FindWindowEx(iIE_Handl_7, 0, "ActXBroadWinBwviewWClass", "Advantech View 001 - main:untitled");
                    break;
                case "CHS":
                    iIE_Handl = tpc.F_FindWindow("IEFrame", "节点 : CTestSCADA - main:untitled"); // 注意是CTestSCADA而不是TestSCADA
                    iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
                    iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "节点 : CTestSCADA - Internet Explorer");  // 注意是CTestSCADA而不是TestSCADA
                    iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
                    iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");
                    iIE_Handl_6 = tpc.F_FindWindowEx(iIE_Handl_5, 0, "AfxOleControl42s", "");
                    iIE_Handl_7 = tpc.F_FindWindowEx(iIE_Handl_6, 0, "AfxWnd42s", "");
                    iWA_MainPage = tpc.F_FindWindowEx(iIE_Handl_7, 0, "ActXBroadWinBwviewWClass", "Advantech View 001 - main:untitled");
                    break;
                case "JPN":
                    iIE_Handl = tpc.F_FindWindow("IEFrame", "ﾉｰﾄﾞ : CTestSCADA - main:untitled"); // 注意是CTestSCADA而不是TestSCADA
                    iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
                    iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "ﾉｰﾄﾞ : CTestSCADA - Internet Explorer");  // 注意是CTestSCADA而不是TestSCADA
                    iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
                    iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");
                    iIE_Handl_6 = tpc.F_FindWindowEx(iIE_Handl_5, 0, "AfxOleControl42s", "");
                    iIE_Handl_7 = tpc.F_FindWindowEx(iIE_Handl_6, 0, "AfxWnd42s", "");
                    iWA_MainPage = tpc.F_FindWindowEx(iIE_Handl_7, 0, "ActXBroadWinBwviewWClass", "Advantech View 001 - main:untitled");
                    break;
                case "KRN":
                case "FRN":

                default:
                    iIE_Handl = tpc.F_FindWindow("IEFrame", "Node : CTestSCADA - main:untitled");
                    iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
                    iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "Node : CTestSCADA - Internet Explorer");
                    iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
                    iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");
                    iIE_Handl_6 = tpc.F_FindWindowEx(iIE_Handl_5, 0, "AfxOleControl42s", "");
                    iIE_Handl_7 = tpc.F_FindWindowEx(iIE_Handl_6, 0, "AfxWnd42s", "");
                    iWA_MainPage = tpc.F_FindWindowEx(iIE_Handl_7, 0, "ActXBroadWinBwviewWClass", "Advantech View 001 - main:untitled");
                    break;
            }
            //int iIE_Handl = tpc.F_FindWindow("IEFrame", "Node : CTestSCADA - main:untitled");   // 注意是CTestSCADA而不是TestSCADA
            //int iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
            //int iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "Node : CTestSCADA - Internet Explorer");    // 注意是CTestSCADA而不是TestSCADA
            //int iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
            //int iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");
            //int iIE_Handl_6 = tpc.F_FindWindowEx(iIE_Handl_5, 0, "AfxOleControl42s", "");
            //int iIE_Handl_7 = tpc.F_FindWindowEx(iIE_Handl_6, 0, "AfxWnd42s", "");
            //int iWA_MainPage = tpc.F_FindWindowEx(iIE_Handl_7, 0, "ActXBroadWinBwviewWClass", "Advantech View 001 - main:untitled");

            if (iWA_MainPage > 0)
            {
                Thread.Sleep(5000);
                tpc.F_PostMessage(iWA_MainPage, tpc.V_WM_KEYDOWN, tpc.V_VK_ESCAPE, 0);
                System.Threading.Thread.Sleep(1000);
            }
            else
                EventLog.AddLog("Cannot get Start View WebAccess Main Page handle");

            // Login keyboard
            //int iLoginKeyboard_Handle = FindWindow("#32770 (Dialog)", "Login");
            EventLog.AddLog("<CloudPC> Login");
            int iLoginKeyboard_Handle;
            switch (slanguage)
            {
                case "ENG":
                    iLoginKeyboard_Handle = tpc.F_FindWindow("#32770", "Login");
                    break;
                case "CHT":
                    iLoginKeyboard_Handle = tpc.F_FindWindow("#32770", "登入");
                    break;
                case "CHS":
                    iLoginKeyboard_Handle = tpc.F_FindWindow("#32770", "登录");
                    break;
                case "JPN":
                    iLoginKeyboard_Handle = tpc.F_FindWindow("#32770", "ﾛｸﾞｲﾝ");
                    break;
                case "KRN":
                case "FRN":

                default:
                    iLoginKeyboard_Handle = tpc.F_FindWindow("#32770", "Login");
                    break;
            }
            int iEnterText = tpc.F_FindWindowEx(iLoginKeyboard_Handle, 0, "Edit", "");
            if (iEnterText > 0)
            {
                //tpc.F_PostMessage(iEnterText, tpc.V_WM_CHAR, 'a', 0);
                //System.Threading.Thread.Sleep(100);
                //tpc.F_PostMessage(iEnterText, tpc.V_WM_CHAR, 'd', 0); //d
                //System.Threading.Thread.Sleep(100);
                //tpc.F_PostMessage(iEnterText, tpc.V_WM_CHAR, 'm', 0); //m
                //System.Threading.Thread.Sleep(100);
                //tpc.F_PostMessage(iEnterText, tpc.V_WM_CHAR, 'i', 0); //i
                //System.Threading.Thread.Sleep(100);
                //tpc.F_PostMessage(iEnterText, tpc.V_WM_CHAR, 'n', 0); //n
                //System.Threading.Thread.Sleep(100);
                SendCharToHandle(iEnterText, 100, "admin");
                tpc.F_PostMessage(iEnterText, tpc.V_WM_KEYDOWN, tpc.V_VK_RETURN, 0);
            }
            else
                EventLog.AddLog("Cannot get Login keyboard handle");

            Thread.Sleep(1000);
            SendKeys.SendWait("^{F5}");
            Thread.Sleep(1000);

            EventLog.AddLog("<CloudPC> Open Point Info window");
            int iPointInfo_Handle;
            switch (slanguage)
            {
                case "ENG":
                    iPointInfo_Handle = tpc.F_FindWindow("#32770", "Point Info");
                    break;
                case "CHT":
                    iPointInfo_Handle = tpc.F_FindWindow("#32770", "點資訊");
                    break;
                case "CHS":
                    iPointInfo_Handle = tpc.F_FindWindow("#32770", "点信息");
                    break;
                case "JPN":
                    iPointInfo_Handle = tpc.F_FindWindow("#32770", "ﾎﾟｲﾝﾄ情報");
                    break;
                case "KRN":
                case "FRN":

                default:
                    iPointInfo_Handle = tpc.F_FindWindow("#32770", "Point Info");
                    break;
            }
            //int iPointInfo_Handle = tpc.F_FindWindow("#32770", "Point Info");
            int iEnterText_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Edit", "");
            if (iEnterText_PointInfo > 0)
            {
                // Acc_0100
                EventLog.AddLog("<CloudPC> Get Acc_0100 point info");
                SendCharToHandle(iEnterText_PointInfo, 100, "Acc_0100");
                System.Threading.Thread.Sleep(1000);
                
                PrintScreen("PlugandPlay_TagInfoSyncTest_Acc_0100", sTestLogFolder);
                for (int i = 1; i <= 10; i++)
                {
                    tpc.F_PostMessage(iEnterText_PointInfo, tpc.V_WM_KEYDOWN, tpc.V_VK_BACK, 0);
                    System.Threading.Thread.Sleep(100);
                }

                // AT_AI0100
                EventLog.AddLog("<CloudPC> Get AT_AI0100 point info");
                SendCharToHandle(iEnterText_PointInfo, 100, "AT_AI0100");
                System.Threading.Thread.Sleep(1000);

                PrintScreen("PlugandPlay_TagInfoSyncTest_AT_AI0100", sTestLogFolder);
                for (int i = 1; i <= 10; i++ )
                {
                    tpc.F_PostMessage(iEnterText_PointInfo, tpc.V_WM_KEYDOWN, tpc.V_VK_BACK, 0);
                    System.Threading.Thread.Sleep(100);
                }

                // AT_AO0100
                EventLog.AddLog("<CloudPC> Get AT_AO0100 point info");
                SendCharToHandle(iEnterText_PointInfo, 100, "AT_AO0100");
                System.Threading.Thread.Sleep(1000);

                PrintScreen("PlugandPlay_TagInfoSyncTest_AT_AO0100", sTestLogFolder);
                for (int i = 1; i <= 10; i++)
                {
                    tpc.F_PostMessage(iEnterText_PointInfo, tpc.V_WM_KEYDOWN, tpc.V_VK_BACK, 0);
                    System.Threading.Thread.Sleep(100);
                }

                // AT_DI0100
                EventLog.AddLog("<CloudPC> Get AT_DI0100 point info");
                SendCharToHandle(iEnterText_PointInfo, 100, "AT_DI0100");
                System.Threading.Thread.Sleep(1000);

                PrintScreen("PlugandPlay_TagInfoSyncTest_AT_DI0100", sTestLogFolder);
                for (int i = 1; i <= 10; i++)
                {
                    tpc.F_PostMessage(iEnterText_PointInfo, tpc.V_WM_KEYDOWN, tpc.V_VK_BACK, 0);
                    System.Threading.Thread.Sleep(100);
                }

                // AT_DO0100
                EventLog.AddLog("<CloudPC> Get AT_DO0100 point info");
                SendCharToHandle(iEnterText_PointInfo, 100, "AT_DO0100");
                System.Threading.Thread.Sleep(1000);

                PrintScreen("PlugandPlay_TagInfoSyncTest_AT_DO0100", sTestLogFolder);
                for (int i = 1; i <= 10; i++)
                {
                    tpc.F_PostMessage(iEnterText_PointInfo, tpc.V_WM_KEYDOWN, tpc.V_VK_BACK, 0);
                    System.Threading.Thread.Sleep(100);
                }

                // Calc_System
                EventLog.AddLog("<CloudPC> Get Calc_System point info");
                SendCharToHandle(iEnterText_PointInfo, 100, "Calc_System");
                System.Threading.Thread.Sleep(1000);

                PrintScreen("PlugandPlay_TagInfoSyncTest_Calc_System", sTestLogFolder);
                for (int i = 1; i <= 15; i++)
                {
                    tpc.F_PostMessage(iEnterText_PointInfo, tpc.V_WM_KEYDOWN, tpc.V_VK_BACK, 0);
                    System.Threading.Thread.Sleep(100);
                }

                // ConAna_0100
                EventLog.AddLog("<CloudPC> Get ConAna_0100 point info");
                SendCharToHandle(iEnterText_PointInfo, 100, "ConAna_0100");
                System.Threading.Thread.Sleep(1000);

                PrintScreen("PlugandPlay_TagInfoSyncTest_ConAna_0100", sTestLogFolder);
                for (int i = 1; i <= 15; i++)
                {
                    tpc.F_PostMessage(iEnterText_PointInfo, tpc.V_WM_KEYDOWN, tpc.V_VK_BACK, 0);
                    System.Threading.Thread.Sleep(100);
                }

                // ConDis_0100
                EventLog.AddLog("<CloudPC> Get ConDis_0100 point info");
                SendCharToHandle(iEnterText_PointInfo, 100, "ConDis_0100");
                System.Threading.Thread.Sleep(1000);

                PrintScreen("PlugandPlay_TagInfoSyncTest_ConDis_0100", sTestLogFolder);
                for (int i = 1; i <= 15; i++)
                {
                    tpc.F_PostMessage(iEnterText_PointInfo, tpc.V_WM_KEYDOWN, tpc.V_VK_BACK, 0);
                    System.Threading.Thread.Sleep(100);
                }

                // OPCDA_0100
                EventLog.AddLog("<CloudPC> Get OPCDA_0100 point info");
                SendCharToHandle(iEnterText_PointInfo, 100, "OPCDA_0100");
                System.Threading.Thread.Sleep(1000);

                PrintScreen("PlugandPlay_TagInfoSyncTest_OPCDA_0100", sTestLogFolder);
                for (int i = 1; i <= 15; i++)
                {
                    tpc.F_PostMessage(iEnterText_PointInfo, tpc.V_WM_KEYDOWN, tpc.V_VK_BACK, 0);
                    System.Threading.Thread.Sleep(100);
                }

                // OPCUA_0100
                EventLog.AddLog("<CloudPC> Get OPCUA_0100 point info");
                SendCharToHandle(iEnterText_PointInfo, 100, "OPCUA_0100");
                System.Threading.Thread.Sleep(1000);

                PrintScreen("PlugandPlay_TagInfoSyncTest_OPCUA_0100", sTestLogFolder);
                for (int i = 1; i <= 15; i++)
                {
                    tpc.F_PostMessage(iEnterText_PointInfo, tpc.V_WM_KEYDOWN, tpc.V_VK_BACK, 0);
                    System.Threading.Thread.Sleep(100);
                }

                // SystemSec_0100
                EventLog.AddLog("<CloudPC> Get SystemSec_0100 point info");
                SendCharToHandle(iEnterText_PointInfo, 100, "SystemSec_0100");
                System.Threading.Thread.Sleep(1000);

                PrintScreen("PlugandPlay_TagInfoSyncTest_SystemSec_0100", sTestLogFolder);
                for (int i = 1; i <= 15; i++)
                {
                    tpc.F_PostMessage(iEnterText_PointInfo, tpc.V_WM_KEYDOWN, tpc.V_VK_BACK, 0);
                    System.Threading.Thread.Sleep(100);
                }
            }
            else
                EventLog.AddLog("Cannot get EnterText_PointInfo handle");

            EventLog.AddLog("Exit Point Info window by press ESC key");
            tpc.F_PostMessage(iPointInfo_Handle, tpc.V_WM_KEYDOWN, tpc.V_VK_ESCAPE, 0);
            api2.Quit();
            PrintStep(api2, "<CloudPC> Quit browser");
        }

        private void SendCharToHandle(int iHandle, int iDelay, string sText)
        {
            var chars = sText.ToCharArray();
            for (int ctr = 0; ctr < chars.Length; ctr++)
            {
                tpc.F_PostMessage(iHandle, tpc.V_WM_CHAR, chars[ctr], 0);
                Thread.Sleep(iDelay);
            }
        }

        private void StartNode(IAdvSeleniumAPI api, string sTestLogFolder)
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
            PrintScreen("Start Node result", sTestLogFolder);
            api.Close();
            EventLog.AddLog("Close start node window and switch to main window");
            api.SwitchToWinHandle(main);        // switch back to original window

            PrintStep(api, "Start Kernel");
        }

        private void PrintScreen(string sFileName, string sFilePath)
        {
            Bitmap myImage = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics g = Graphics.FromImage(myImage);
            g.CopyFromScreen(new Point(0, 0), new Point(0, 0), new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height));
            IntPtr dc1 = g.GetHdc();
            g.ReleaseHdc(dc1);
            //myImage.Save(@"c:\screen0.jpg");
            myImage.Save(string.Format("{0}\\{1}_{2:yyyyMMdd_hhmmss}.jpg", sFilePath, sFileName, DateTime.Now));
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

        private void ReturnSCADAPage(IAdvSeleniumAPI api)
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMainRight.asp') and contains(@href, 'name=TestSCADA')]").Click();

        }

        private void PrintStep(IAdvSeleniumAPI api, string sTestItem)
        {
            DataGridViewRow dgvRow;
            DataGridViewCell dgvCell;

            var list = api.GetStepResult();
            foreach (var item in list)
            {
                AdvSeleniumAPI.ResultClass _res = (AdvSeleniumAPI.ResultClass)item;
                //
                dgvRow = new DataGridViewRow();
                if (_res.Res == "fail")
                    dgvRow.DefaultCellStyle.ForeColor = Color.Red;
                dgvCell = new DataGridViewTextBoxCell(); //Column Time
                //
                if (_res == null) continue;
                //
                dgvCell.Value = sTestItem;
                dgvRow.Cells.Add(dgvCell);
                //
                dgvCell = new DataGridViewTextBoxCell();
                dgvCell.Value = _res.Decp;
                dgvRow.Cells.Add(dgvCell);
                //
                dgvCell = new DataGridViewTextBoxCell();
                dgvCell.Value = _res.Res;
                dgvRow.Cells.Add(dgvCell);
                //
                dgvCell = new DataGridViewTextBoxCell();
                dgvCell.Value = _res.Err;
                dgvRow.Cells.Add(dgvCell);
                //
                dgvCell = new DataGridViewTextBoxCell();
                dgvCell.Value = _res.Tdev;
                dgvRow.Cells.Add(dgvCell);

                m_DataGridViewCtrlAddDataRow(dgvRow);
            }
            Application.DoEvents();
        }

        private void Start_Click(object sender, EventArgs e)
        {
            long lErrorCode = (long)ErrorCode.SUCCESS;
            EventLog.AddLog("===PlugandPlay_TagInfoSyncTest start===");
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address(Ground PC)= " + WebAccessIP.Text);
            EventLog.AddLog("WebAccess IP address(Cloud PC)= " + WebAccessIP2.Text);
            lErrorCode = Form1_Load(ProjectName.Text, ProjectName2.Text, WebAccessIP.Text, WebAccessIP2.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===PlugandPlay_TagInfoSyncTest end===");
        }

        private void InitialRequiredInfo(string sFilePath)
        {
            StringBuilder sDefaultUserLanguage = new StringBuilder(255);
            StringBuilder sDefaultProjectName1 = new StringBuilder(255);
            StringBuilder sDefaultProjectName2 = new StringBuilder(255);
            StringBuilder sDefaultIP1 = new StringBuilder(255);
            StringBuilder sDefaultIP2 = new StringBuilder(255);
            /*
            tpc.F_WritePrivateProfileString("ProjectName", "Ground PC or Primary PC", "TestProject", @"C:\WebAccessAutoTestSetting.ini");
            tpc.F_WritePrivateProfileString("ProjectName", "Cloud PC or Backup PC", "CTestProject", @"C:\WebAccessAutoTestSetting.ini");
            tpc.F_WritePrivateProfileString("IP", "Ground PC or Primary PC", "172.18.3.62", @"C:\WebAccessAutoTestSetting.ini");
            tpc.F_WritePrivateProfileString("IP", "Cloud PC or Backup PC", "172.18.3.65", @"C:\WebAccessAutoTestSetting.ini");
            */
            tpc.F_GetPrivateProfileString("UserInfo", "Language", "NA", sDefaultUserLanguage, 255, sFilePath);
            tpc.F_GetPrivateProfileString("ProjectName", "Ground PC or Primary PC", "NA", sDefaultProjectName1, 255, sFilePath);
            tpc.F_GetPrivateProfileString("ProjectName", "Cloud PC or Backup PC", "NA", sDefaultProjectName2, 255, sFilePath);
            tpc.F_GetPrivateProfileString("IP", "Ground PC or Primary PC", "NA", sDefaultIP1, 255, sFilePath);
            tpc.F_GetPrivateProfileString("IP", "Cloud PC or Backup PC", "NA", sDefaultIP2, 255, sFilePath);
            slanguage = sDefaultUserLanguage.ToString();    // 在這邊讀取使用語言

            ProjectName.Text = sDefaultProjectName1.ToString();
            WebAccessIP.Text = sDefaultIP1.ToString();

            ProjectName2.Text = sDefaultProjectName2.ToString();
            WebAccessIP2.Text = sDefaultIP2.ToString();
        }

        private void CheckifIniFileChange()
        {
            StringBuilder sDefaultProjectName1 = new StringBuilder(255);
            StringBuilder sDefaultProjectName2 = new StringBuilder(255);
            StringBuilder sDefaultIP1 = new StringBuilder(255);
            StringBuilder sDefaultIP2 = new StringBuilder(255);
            if (System.IO.File.Exists(sIniFilePath))    // 比對ini檔與ui上的值是否相同
            {
                EventLog.AddLog(".ini file exist, check if .ini file need to update");
                tpc.F_GetPrivateProfileString("ProjectName", "Ground PC or Primary PC", "NA", sDefaultProjectName1, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("ProjectName", "Cloud PC or Backup PC", "NA", sDefaultProjectName2, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("IP", "Ground PC or Primary PC", "NA", sDefaultIP1, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("IP", "Cloud PC or Backup PC", "NA", sDefaultIP2, 255, sIniFilePath);

                if (ProjectName.Text != sDefaultProjectName1.ToString())
                {
                    tpc.F_WritePrivateProfileString("ProjectName", "Ground PC or Primary PC", ProjectName.Text, sIniFilePath);
                    EventLog.AddLog("New ProjectName update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultProjectName1.ToString());
                    EventLog.AddLog("New ini:" + ProjectName.Text);
                }
                if (WebAccessIP.Text != sDefaultIP1.ToString())
                {
                    tpc.F_WritePrivateProfileString("IP", "Ground PC or Primary PC", WebAccessIP.Text, sIniFilePath);
                    EventLog.AddLog("New WebAccessIP update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultIP1.ToString());
                    EventLog.AddLog("New ini:" + WebAccessIP.Text);
                }
                if (ProjectName2.Text != sDefaultProjectName2.ToString())
                {
                    tpc.F_WritePrivateProfileString("ProjectName", "Cloud PC or Backup PC", ProjectName2.Text, sIniFilePath);
                    EventLog.AddLog("New ProjectName update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultProjectName2.ToString());
                    EventLog.AddLog("New ini:" + ProjectName2.Text);
                }
                if (WebAccessIP2.Text != sDefaultIP2.ToString())
                {
                    tpc.F_WritePrivateProfileString("IP", "Cloud PC or Backup PC", WebAccessIP2.Text, sIniFilePath);
                    EventLog.AddLog("New WebAccessIP update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultIP2.ToString());
                    EventLog.AddLog("New ini:" + WebAccessIP2.Text);
                }
            }
            else
            {
                EventLog.AddLog(".ini file not exist, create new .ini file. Path: " + sIniFilePath);
                tpc.F_WritePrivateProfileString("ProjectName", "Ground PC or Primary PC", ProjectName.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("ProjectName", "Cloud PC or Backup PC", ProjectName2.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("IP", "Ground PC or Primary PC", WebAccessIP.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("IP", "Cloud PC or Backup PC", WebAccessIP2.Text, sIniFilePath);
            }
        }

    }
}
