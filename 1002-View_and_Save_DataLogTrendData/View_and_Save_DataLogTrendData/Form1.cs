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

namespace View_and_Save_DataLogTrendData
{
    public partial class Form1 : Form, iATester.iCom
    {
        IAdvSeleniumAPI api;
        cThirdPartyToolControl tpc = new cThirdPartyToolControl();

        private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
        private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
        internal const int Max_Rows_Val = 65535;
        string baseUrl;
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
            EventLog.AddLog("===View and Save DataLogTrendData start (by iATester)===");
            if (System.IO.File.Exists(sIniFilePath))    // 再load一次
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===View and Save DataLogTrendData end (by iATester)===");

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

        long Form1_Load(string sProjectName, string sWebAccessIP, string sTestLogFolder, string sBrowser)
        {
            baseUrl = "http://" + sWebAccessIP;

            if (sBrowser == "Internet Explorer")
            {
                EventLog.AddLog("Browser= Internet Explorer");
                api = new AdvSeleniumAPI("IE", "");
                System.Threading.Thread.Sleep(1000);
            }
            else if (sBrowser == "Mozilla FireFox")
            {
                EventLog.AddLog("Browser= Mozilla FireFox");
                api = new AdvSeleniumAPI("FireFox", "");
                System.Threading.Thread.Sleep(1000);
            }


            // Launch Firefox and login
            api.LinkWebUI(baseUrl + "/broadWeb/bwconfig.asp?username=admin");
            api.ById("userField").Enter("").Submit().Exe();
            PrintStep("Login WebAccess");

            // Configure project by project name
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
            PrintStep("Configure project");

            EventLog.AddLog("Start view data log trend");
            StartViewDataLogTrend(sProjectName, sTestLogFolder);

            Thread.Sleep(500);
            api.Quit();
            PrintStep("Quit browser");

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

        private void StartViewDataLogTrend(string sProjectName, string sTestLogFolder)
        {
            // Start view
            EventLog.AddLog("Start view");
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//tr[2]/td/a/font").Click();
            PrintStep("Start View");

            Thread.Sleep(5000);

            // Control browser
            EventLog.AddLog("Control browser");
            int iIE_Handl, iIE_Handl_2, iIE_Handl_3, iIE_Handl_4, iIE_Handl_5, iIE_Handl_6, iIE_Handl_7, iWA_MainPage = 0;
            switch (slanguage)
            {
                case "ENG":
                    iIE_Handl = tpc.F_FindWindow("IEFrame", "Node : TestSCADA - main:untitled");
                    iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
                    iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "Node : TestSCADA - Internet Explorer");
                    iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
                    iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");
                    iIE_Handl_6 = tpc.F_FindWindowEx(iIE_Handl_5, 0, "AfxOleControl42s", "");
                    iIE_Handl_7 = tpc.F_FindWindowEx(iIE_Handl_6, 0, "AfxWnd42s", "");
                    iWA_MainPage = tpc.F_FindWindowEx(iIE_Handl_7, 0, "ActXBroadWinBwviewWClass", "Advantech View 001 - main:untitled");
                    break;
                case "CHT":
                    iIE_Handl = tpc.F_FindWindow("IEFrame", "節點 : TestSCADA - main:untitled");
                    iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
                    iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "節點 : TestSCADA - Internet Explorer");
                    iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
                    iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");
                    iIE_Handl_6 = tpc.F_FindWindowEx(iIE_Handl_5, 0, "AfxOleControl42s", "");
                    iIE_Handl_7 = tpc.F_FindWindowEx(iIE_Handl_6, 0, "AfxWnd42s", "");
                    iWA_MainPage = tpc.F_FindWindowEx(iIE_Handl_7, 0, "ActXBroadWinBwviewWClass", "Advantech View 001 - main:untitled");
                    break;
                case "CHS":
                    iIE_Handl = tpc.F_FindWindow("IEFrame", "节点 : TestSCADA - main:untitled");
                    iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
                    iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "节点 : TestSCADA - Internet Explorer");
                    iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
                    iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");
                    iIE_Handl_6 = tpc.F_FindWindowEx(iIE_Handl_5, 0, "AfxOleControl42s", "");
                    iIE_Handl_7 = tpc.F_FindWindowEx(iIE_Handl_6, 0, "AfxWnd42s", "");
                    iWA_MainPage = tpc.F_FindWindowEx(iIE_Handl_7, 0, "ActXBroadWinBwviewWClass", "Advantech View 001 - main:untitled");
                    break;
                case "JPN":
                    iIE_Handl = tpc.F_FindWindow("IEFrame", "ﾉｰﾄﾞ : TestSCADA - main:untitled");
                    iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
                    iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "ﾉｰﾄﾞ : TestSCADA - Internet Explorer");
                    iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
                    iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");
                    iIE_Handl_6 = tpc.F_FindWindowEx(iIE_Handl_5, 0, "AfxOleControl42s", "");
                    iIE_Handl_7 = tpc.F_FindWindowEx(iIE_Handl_6, 0, "AfxWnd42s", "");
                    iWA_MainPage = tpc.F_FindWindowEx(iIE_Handl_7, 0, "ActXBroadWinBwviewWClass", "Advantech View 001 - main:untitled");
                    break;
                case "KRN":
                    iIE_Handl = tpc.F_FindWindow("IEFrame", "노드 : TestSCADA - main:untitled");
                    iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
                    iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "노드 : TestSCADA - Internet Explorer");
                    iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
                    iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");
                    iIE_Handl_6 = tpc.F_FindWindowEx(iIE_Handl_5, 0, "AfxOleControl42s", "");
                    iIE_Handl_7 = tpc.F_FindWindowEx(iIE_Handl_6, 0, "AfxWnd42s", "");
                    iWA_MainPage = tpc.F_FindWindowEx(iIE_Handl_7, 0, "ActXBroadWinBwviewWClass", "Advantech View 001 - main:untitled");
                    break;
                case "FRN":
                    iIE_Handl = tpc.F_FindWindow("IEFrame", "Noeud : TestSCADA - main:untitled");
                    iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
                    iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "Noeud : TestSCADA - Internet Explorer");
                    iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
                    iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");
                    iIE_Handl_6 = tpc.F_FindWindowEx(iIE_Handl_5, 0, "AfxOleControl42s", "");
                    iIE_Handl_7 = tpc.F_FindWindowEx(iIE_Handl_6, 0, "AfxWnd42s", "");
                    iWA_MainPage = tpc.F_FindWindowEx(iIE_Handl_7, 0, "ActXBroadWinBwviewWClass", "Advantech View 001 - main:untitled");
                    break;

                default:
                    iIE_Handl = tpc.F_FindWindow("IEFrame", "Node : TestSCADA - main:untitled");
                    iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
                    iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "Node : TestSCADA - Internet Explorer");
                    iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
                    iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");
                    iIE_Handl_6 = tpc.F_FindWindowEx(iIE_Handl_5, 0, "AfxOleControl42s", "");
                    iIE_Handl_7 = tpc.F_FindWindowEx(iIE_Handl_6, 0, "AfxWnd42s", "");
                    iWA_MainPage = tpc.F_FindWindowEx(iIE_Handl_7, 0, "ActXBroadWinBwviewWClass", "Advantech View 001 - main:untitled");
                    break;
            }
            /*
            int iIE_Handl = tpc.F_FindWindow("IEFrame", "Node : TestSCADA - main:untitled");
            int iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
            int iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "Node : TestSCADA - Internet Explorer");
            int iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
            int iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");
            int iIE_Handl_6 = tpc.F_FindWindowEx(iIE_Handl_5, 0, "AfxOleControl42s", "");
            int iIE_Handl_7 = tpc.F_FindWindowEx(iIE_Handl_6, 0, "AfxWnd42s", "");
            int iWA_MainPage = tpc.F_FindWindowEx(iIE_Handl_7, 0, "ActXBroadWinBwviewWClass", "Advantech View 001 - main:untitled");
            */

            if (iWA_MainPage > 0)
            {
                //SendMessage(iWA_MainPage, BM_CLICK, 0, 0);
                //SendMessage(iWA_MainPage, WM_RBUTTONDOWN, 0, 0);
                //SendMessage(iWA_MainPage, WM_RBUTTONDOWN, MK_RBUTTON, 0);
                tpc.F_PostMessage(iWA_MainPage, tpc.V_WM_KEYDOWN, tpc.V_VK_ESCAPE, 0);
                System.Threading.Thread.Sleep(1500);
            }
            else
                EventLog.AddLog("Cannot get Start View WebAccess Main Page handle");

            // Login keyboard
            EventLog.AddLog("admin login");
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
                    iLoginKeyboard_Handle = tpc.F_FindWindow("#32770", "로그인");
                    break;
                case "FRN":
                    iLoginKeyboard_Handle = tpc.F_FindWindow("#32770", "Connexion");
                    break;

                default:
                    iLoginKeyboard_Handle = tpc.F_FindWindow("#32770", "Login");
                    break;
            }

            int iEnterText = tpc.F_FindWindowEx(iLoginKeyboard_Handle, 0, "Edit", "");
            if (iEnterText > 0)
            {
                tpc.F_PostMessage(iEnterText, tpc.V_WM_CHAR, 'a', 0);
                System.Threading.Thread.Sleep(100);
                tpc.F_PostMessage(iEnterText, tpc.V_WM_CHAR, 'd', 0); //d
                System.Threading.Thread.Sleep(100);
                tpc.F_PostMessage(iEnterText, tpc.V_WM_CHAR, 'm', 0); //m
                System.Threading.Thread.Sleep(100);
                tpc.F_PostMessage(iEnterText, tpc.V_WM_CHAR, 'i', 0); //i
                System.Threading.Thread.Sleep(100);
                tpc.F_PostMessage(iEnterText, tpc.V_WM_CHAR, 'n', 0); //n
                System.Threading.Thread.Sleep(100);
                tpc.F_PostMessage(iEnterText, tpc.V_WM_KEYDOWN, tpc.V_VK_RETURN, 0);
                System.Threading.Thread.Sleep(500);
                //PostMessage(iWA_MainPage, WM_KEYDOWN, VK_F4, 0);
                SendKeys.SendWait("{F4}");
                System.Threading.Thread.Sleep(1000);
            }
            else
                EventLog.AddLog("Cannot get Login keyboard handle");

            int iDataLogTrend_Handle;
            switch (slanguage)
            {
                case "ENG":
                    iDataLogTrend_Handle = tpc.F_FindWindow("#32770", "Datalog Trend List");
                    break;
                case "CHT":
                    iDataLogTrend_Handle = tpc.F_FindWindow("#32770", "資料記錄趨勢列表");
                    break;
                case "CHS":
                    iDataLogTrend_Handle = tpc.F_FindWindow("#32770", "历史趋势列表");
                    break;
                case "JPN":
                    iDataLogTrend_Handle = tpc.F_FindWindow("#32770", "ﾃﾞｰﾀﾛｸﾞ ﾄﾚﾝﾄﾞ一覧");
                    break;
                case "KRN":
                    iDataLogTrend_Handle = tpc.F_FindWindow("#32770", "데이터로그 트랜드 리스트");
                    break;
                case "FRN":
                    iDataLogTrend_Handle = tpc.F_FindWindow("#32770", "Liste courbes historiques");
                    break;

                default:
                    iDataLogTrend_Handle = tpc.F_FindWindow("#32770", "Datalog Trend List");
                    break;
            }
            int iEnterText2 = tpc.F_FindWindowEx(iDataLogTrend_Handle, 0, "Edit", "");
            if (iEnterText2 > 0)
                tpc.F_PostMessage(iEnterText2, tpc.V_WM_KEYDOWN, tpc.V_VK_RETURN, 0);
            else
                EventLog.AddLog("Cannot get DataLog Trend List handle");
            /*
            int iOK_Button_of_DataLogTrendList = tpc.F_FindWindowEx(iDataLogTrend_Handle, 0, "Button", "OK");
            if (iOK_Button_of_DataLogTrendList > 0)
            {
                // Change interval value
                EventLog.AddLog("Click Change button of PointInfo window");
                tpc.F_PostMessage(iOK_Button_of_DataLogTrendList, tpc.V_BM_CLICK, 0, 0);
                System.Threading.Thread.Sleep(1000);
            }
            else
                EventLog.AddLog("Cannot get iOK_Button_of_DataLogTrendList handle");

            Thread.Sleep(3000);
            PrintScreen("DataLogData", sTestLogFolder);
            */
            for (int iInterval = 1; iInterval <= 3; iInterval++)
            {
                string[] sInterval = { "1 sec", "2 sec", "5 sec", "10 sec", "20 sec", "40 sec", 
                                       "1Min", "2Min", "3Min", "4Min", "8Min", "16Min", "28Min", 
                                       "1Hour", "2Hour", "4Hour", "6Hour", "8Hour", "12Hour", 
                                       "1Day", "2Day", "3Day", "4Day", "5Day"};
                int iii = sInterval.Length;
                Thread.Sleep(4000);
                SendKeys.SendWait("{F4}"); // Right most
                Thread.Sleep(4000);

                PrintScreen(string.Format("DataLogTrend_Interval_{0}_RightMost_Step1", sInterval[iInterval-1]), sTestLogFolder);

                SendKeys.SendWait("{F2}"); // Left
                Thread.Sleep(4000);

                PrintScreen(string.Format("DataLogTrend_Interval_{0}_Left_Step2", sInterval[iInterval - 1]), sTestLogFolder);

                SendKeys.SendWait("{F2}"); // Left
                Thread.Sleep(4000);

                PrintScreen(string.Format("DataLogTrend_Interval_{0}_Right_Step3", sInterval[iInterval - 1]), sTestLogFolder);

                SendKeys.SendWait("{F3}"); // Right
                Thread.Sleep(4000);

                PrintScreen(string.Format("DataLogTrend_Interval_{0}_Right_Step4", sInterval[iInterval - 1]), sTestLogFolder);

                SendKeys.SendWait("{F3}"); // Right
                Thread.Sleep(4000);

                PrintScreen(string.Format("DataLogTrend_Interval_{0}_Right_Step5", sInterval[iInterval - 1]), sTestLogFolder);

                SendKeys.SendWait("{F1}"); // Left most
                Thread.Sleep(4000);

                PrintScreen(string.Format("DataLogTrend_Interval_{0}_LeftMost_Step6", sInterval[iInterval - 1]), sTestLogFolder);

                SendKeys.SendWait("{F4}"); // Right most
                Thread.Sleep(4000);

                PrintScreen(string.Format("DataLogTrend_Interval_{0}_RightMost_Step7", sInterval[iInterval - 1]), sTestLogFolder);

                Thread.Sleep(1000);
                SendKeys.SendWait("^{F5}"); // quick key of PointInfo window
                Thread.Sleep(1000);

                EventLog.AddLog("Open Point Info window");
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
                        iPointInfo_Handle = tpc.F_FindWindow("#32770", "포인트 정보");
                        break;
                    case "FRN":
                        iPointInfo_Handle = tpc.F_FindWindow("#32770", "Info point");
                        break;

                    default:
                        iPointInfo_Handle = tpc.F_FindWindow("#32770", "Point Info");
                        break;
                }

                //  用來改變interval
                //int iPointInfo_Handle = tpc.F_FindWindow("#32770", "Point Info");
                int iEnterText_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Edit", "");
                if (iEnterText_PointInfo > 0)
                {
                    EventLog.AddLog("Change Datalog Trend Interval");
                    SendCharToHandle(iEnterText_PointInfo, 100, "%ADTRDST");
                    System.Threading.Thread.Sleep(1000);
                }
                else
                    EventLog.AddLog("Cannot get EnterText_PointInfo handle");

                
                int iChange_Button_of_PointInfo=0;
                switch (slanguage)
                {
                    case "ENG":
                        iChange_Button_of_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "Change");
                        break;
                    case "CHT":
                        iChange_Button_of_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "改變");
                        break;
                    case "CHS":
                        iChange_Button_of_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "改变");
                        break;
                    case "JPN":
                        iChange_Button_of_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "変更");
                        break;
                    case "KRN":
                        iChange_Button_of_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "변경");
                        break;
                    case "FRN":
                        iChange_Button_of_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "Modifier");
                        break;

                    default:
                        iChange_Button_of_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "Change");
                        break;
                }

                if (iChange_Button_of_PointInfo > 0)
                {
                    // Change interval value
                    EventLog.AddLog("Click Change button of PointInfo window");
                    tpc.F_PostMessage(iChange_Button_of_PointInfo, tpc.V_BM_CLICK, 0, 0);
                    System.Threading.Thread.Sleep(1000);
                }
                else
                    EventLog.AddLog("Cannot get iChange_Button_of_PointInfo handle");

                int iEditWindow_of_interval = tpc.F_FindWindow("#32770", "%ADTRDST");
                int iEnterValue_of_interval = tpc.F_FindWindowEx(iEditWindow_of_interval, 0, "Button", "BW_DSPINUP");      // 換設定interval視窗裡的向右鍵
                if (iEnterValue_of_interval > 0)
                {
                    // Change interval value
                    EventLog.AddLog("Click add one interval button Change window");
                    tpc.F_PostMessage(iEnterValue_of_interval, tpc.V_BM_CLICK, 0, 0);
                    System.Threading.Thread.Sleep(1000);
                }
                else
                    EventLog.AddLog("Cannot get iEnterValue_of_interval handle");

                
                int iEnterButton_ChangeWindow = 0;
                switch (slanguage)
                {
                    case "ENG":
                        iEnterButton_ChangeWindow = tpc.F_FindWindowEx(iEditWindow_of_interval, 0, "Button", "Enter");   // 按設定interval視窗裡的enter
                        break;
                    case "CHT":
                        iEnterButton_ChangeWindow = tpc.F_FindWindowEx(iEditWindow_of_interval, 0, "Button", "輸入");
                        break;
                    case "CHS":
                        iEnterButton_ChangeWindow = tpc.F_FindWindowEx(iEditWindow_of_interval, 0, "Button", "回车");
                        break;
                    case "JPN":
                        iEnterButton_ChangeWindow = tpc.F_FindWindowEx(iEditWindow_of_interval, 0, "Button", "ｴﾝﾀｰ");
                        break;
                    case "KRN":
                        iEnterButton_ChangeWindow = tpc.F_FindWindowEx(iEditWindow_of_interval, 0, "Button", "Enter");
                        break;
                    case "FRN":
                        iEnterButton_ChangeWindow = tpc.F_FindWindowEx(iEditWindow_of_interval, 0, "Button", "Valider");
                        break;

                    default:
                        iEnterButton_ChangeWindow = tpc.F_FindWindowEx(iEditWindow_of_interval, 0, "Button", "Enter");   // 按設定interval視窗裡的enter
                        break;
                }

                if (iEnterButton_ChangeWindow > 0)
                {
                    // Change interval value
                    EventLog.AddLog("Click Enter button Change window");
                    tpc.F_PostMessage(iEnterButton_ChangeWindow, tpc.V_BM_CLICK, 0, 0);
                    System.Threading.Thread.Sleep(1000);
                }
                else
                    EventLog.AddLog("Cannot get iEnterButton_ChangeWindow handle");

                int iExitButton_PointInfo = 0;
                switch (slanguage)
                {
                    case "ENG":
                        iExitButton_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "Exit");    // 按PointInfo視窗裡的Exit
                        break;
                    case "CHT":
                        iExitButton_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "退出");
                        break;
                    case "CHS":
                        iExitButton_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "退出");
                        break;
                    case "JPN":
                        iExitButton_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "終了");
                        break;
                    case "KRN":
                        iExitButton_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "종료");
                        break;
                    case "FRN":
                        iExitButton_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "Quitter");
                        break;

                    default:
                        iExitButton_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "Exit");    // 按PointInfo視窗裡的Exit
                        break;
                }

                if (iExitButton_PointInfo > 0)
                {
                    // Change interval value
                    EventLog.AddLog("Click Exit button of PointInfo window");
                    tpc.F_PostMessage(iExitButton_PointInfo, tpc.V_BM_CLICK, 0, 0);
                    System.Threading.Thread.Sleep(1000);
                }
                else
                    EventLog.AddLog("Cannot get iExitButton_PointInfo handle");

                if(iInterval == sInterval.Length)
                {
                    Thread.Sleep(4000);
                    SendKeys.SendWait("{F4}"); // Right most
                    Thread.Sleep(4000);

                    PrintScreen(string.Format("DataLogTrend_Interval_{0}_RightMost_Step1", sInterval[iInterval - 1]), sTestLogFolder);

                    SendKeys.SendWait("{F2}"); // Left
                    Thread.Sleep(4000);

                    PrintScreen(string.Format("DataLogTrend_Interval_{0}_Left_Step2", sInterval[iInterval - 1]), sTestLogFolder);

                    SendKeys.SendWait("{F2}"); // Left
                    Thread.Sleep(4000);

                    PrintScreen(string.Format("DataLogTrend_Interval_{0}_Right_Step3", sInterval[iInterval - 1]), sTestLogFolder);

                    SendKeys.SendWait("{F3}"); // Right
                    Thread.Sleep(4000);

                    PrintScreen(string.Format("DataLogTrend_Interval_{0}_Right_Step4", sInterval[iInterval - 1]), sTestLogFolder);

                    SendKeys.SendWait("{F3}"); // Right
                    Thread.Sleep(4000);

                    PrintScreen(string.Format("DataLogTrend_Interval_{0}_Right_Step5", sInterval[iInterval - 1]), sTestLogFolder);

                    SendKeys.SendWait("{F1}"); // Left most
                    Thread.Sleep(4000);

                    PrintScreen(string.Format("DataLogTrend_Interval_{0}_LeftMost_Step6", sInterval[iInterval - 1]), sTestLogFolder);

                    SendKeys.SendWait("{F4}"); // Right most
                    Thread.Sleep(4000);

                    PrintScreen(string.Format("DataLogTrend_Interval_{0}_RightMost_Step7", sInterval[iInterval - 1]), sTestLogFolder);
                }

                /*
                ////////這邊開始左左右右////////
                DateTime dt = DateTime.Now.AddSeconds(-270); // 開始時間設定為4分半前
                for (int LLRR = 1; LLRR <= 4; LLRR++)
                {
                    Thread.Sleep(1000);
                    SendKeys.SendWait("^{F5}"); // quick key of PointInfo window
                    Thread.Sleep(1000);

                    EventLog.AddLog("Open Point Info window");
                    //int iPointInfo_Handle;
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
                            iPointInfo_Handle = tpc.F_FindWindow("#32770", "포인트 정보");
                            break;
                        case "FRN":
                            iPointInfo_Handle = tpc.F_FindWindow("#32770", "Info point");
                            break;

                        default:
                            iPointInfo_Handle = tpc.F_FindWindow("#32770", "Point Info");
                            break;
                    }

                    //用來改變時間
                    iEnterText_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Edit", "");
                    if (iEnterText_PointInfo > 0)
                    {
                        EventLog.AddLog("Change Datalog Trend Interval");
                        SendCharToHandle(iEnterText_PointInfo, 100, "%TDTRDSTM");
                        System.Threading.Thread.Sleep(1000);
                    }
                    else
                        EventLog.AddLog("Cannot get EnterText_PointInfo handle");
                                                                                                                    // 注意! 尚未加上多國語言版本
                    iChange_Button_of_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "Change");     // 按 PointInfo視窗裡的change 按扭
                    if (iChange_Button_of_PointInfo > 0)
                    {
                        // Change interval value
                        EventLog.AddLog("Click Change button of PointInfo window");
                        tpc.F_PostMessage(iChange_Button_of_PointInfo, tpc.V_BM_CLICK, 0, 0);
                        System.Threading.Thread.Sleep(1000);
                    }
                    else
                        EventLog.AddLog("Cannot get iChange_Button_of_PointInfo handle");

                    int iTDTRDSTM_window = tpc.F_FindWindow("#32770", "%TDTRDSTM");
                    int iEditField_of_TDTRDSTM = tpc.F_FindWindowEx(iTDTRDSTM_window, 0, "Edit", "");
                    int iBackspace_of_TDTRDSTM = tpc.F_FindWindowEx(iTDTRDSTM_window, 0, "Button", "");
                    for (int i = 0; i < 37; i++)    // 換設定interval視窗裡的Backspace
                    {
                        iBackspace_of_TDTRDSTM = tpc.F_FindWindowEx(iTDTRDSTM_window, iBackspace_of_TDTRDSTM, "Button", "");
                    }
                    if (iBackspace_of_TDTRDSTM > 0)
                    {
                        // 
                        EventLog.AddLog("Click Backspace button Change window");
                        tpc.F_PostMessage(iBackspace_of_TDTRDSTM, tpc.V_BM_CLICK, 0, 0);
                        System.Threading.Thread.Sleep(1000);
                    }
                    else
                        EventLog.AddLog("Cannot get iBackspace_of_TDTRDSTM handle");

                    if (iEditField_of_TDTRDSTM > 0)
                    {
                        string sTime;
                        if(LLRR <= 2)   //前兩次往左推2單位時間
                            sTime = string.Format("{0:yyyy/MM/dd HH:mm:ss}", dt.AddSeconds(-(90*LLRR)));
                        else            //後兩次往右推2單位時間
                            sTime = string.Format("{0:yyyy/MM/dd HH:mm:ss}", dt.AddSeconds(90*(LLRR-2)));

                        EventLog.AddLog("Enter the date time in Edit field");
                        SendCharToHandle(iEditField_of_TDTRDSTM, 100, sTime);
                        System.Threading.Thread.Sleep(1000);
                        tpc.F_PostMessage(iEditField_of_TDTRDSTM, tpc.V_WM_KEYDOWN, tpc.V_VK_RETURN, 0);
                    }
                    else
                        EventLog.AddLog("Cannot get iEditField_of_TDTRDSTM handle");

                    Thread.Sleep(1000);
                    iExitButton_PointInfo = tpc.F_FindWindowEx(iPointInfo_Handle, 0, "Button", "Exit");    // 按PointInfo視窗裡的Exit
                                                                                                           // 注意! 尚未加上多國語言版本
                    if (iExitButton_PointInfo > 0)
                    {
                        // Change interval value
                        EventLog.AddLog("Click Exit button of PointInfo window");
                        tpc.F_PostMessage(iExitButton_PointInfo, tpc.V_BM_CLICK, 0, 0);
                        System.Threading.Thread.Sleep(1000);
                    }
                    else
                        EventLog.AddLog("Cannot get iExitButton_PointInfo handle");
                }
                */
            }

            // Export Data (ctrl+F4) , 會在C:\WebAccess\Client\TestProject_TestSCADA 產生 $view001.htm 檔
            EventLog.AddLog("Send ctrl+F4 keyboard command to export data");
            if (iWA_MainPage > 0)
            {
                /*
                SendMessage(iWA_MainPage, WM_KEYDOWN, VK_CONTROL, 0);
                SendMessage(iWA_MainPage, WM_KEYDOWN, VK_F4, 0);
                Thread.Sleep(1000);
                SendMessage(iWA_MainPage, WM_KEYUP, VK_F4, 0);
                SendMessage(iWA_MainPage, WM_KEYUP, VK_CONTROL, 0);
                */
                /*
                keybd_event(VK_CONTROL, 0, 0, 0);
                PostMessage(iWA_MainPage, WM_KEYDOWN, VK_F4, 0);
                PostMessage(iWA_MainPage, WM_KEYUP, VK_F4, 0);
                keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
                */
                Thread.Sleep(1000);
                SendKeys.SendWait("^{F4}"); // [Aaron] I have no idea how to implement ctrl+F4...so this poor method used.
                // this method will press ctrl + F4 key with no handle input
                Thread.Sleep(1500);
            }
            else
                EventLog.AddLog("Cannot get Start View WebAccess Main Page handle");

            // copy $view001.htm to log path.
            EventLog.AddLog("copy $view001.htm to log path.");
            {
                string fileNameSrc = "$view001.htm";
                string fileNameTar = string.Format("DataLogGroup_{0:yyyyMMdd_hhmmss}.htm", DateTime.Now);

                string sourcePath = string.Format(@"C:\WebAccess\Client\{0}_TestSCADA", sProjectName);
                string targetPath = sTestLogFolder;

                // Use Path class to manipulate file and directory paths.
                string sourceFile = System.IO.Path.Combine(sourcePath, fileNameSrc);
                string destFile = System.IO.Path.Combine(targetPath, fileNameTar);

                if (System.IO.File.Exists(sourceFile))
                    System.IO.File.Copy(sourceFile, destFile, true);
                else
                    EventLog.AddLog("Cannot get $view001.htm file");
            }

            if (iWA_MainPage > 0)
            {
                Thread.Sleep(5000); //delay 5s before close $view001 window
                EventLog.AddLog("Send alt+F4 keyboard command to close $view001.htm window");
                SendKeys.SendWait("%{F4}");
                Thread.Sleep(500);
            }
            else
                EventLog.AddLog("Cannot find $view001.htm window");

            //PrintScreen("DataLogData", sTestLogFolder);
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

        private void PrintScreen(string sFileName, string sFilePath)
        {
            Bitmap myImage = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics g = Graphics.FromImage(myImage);
            g.CopyFromScreen(new Point(0, 0), new Point(0, 0), new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height));
            IntPtr dc1 = g.GetHdc();
            g.ReleaseHdc(dc1);
            //myImage.Save(@"c:\screen0.jpg");
            myImage.Save(string.Format("{0}\\{1}_{2:yyyyMMdd_HHmmss}.jpg", sFilePath, sFileName, DateTime.Now));
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

        private void ReturnSCADAPage()
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMainRight.asp') and contains(@href, 'name=TestSCADA')]").Click();

        }

        private void PrintStep(string sTestItem)
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
            EventLog.AddLog("===View and Save DataLogTrendData start===");
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===View and Save DataLogTrendData end===");
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
            }
            else
            {
                EventLog.AddLog(".ini file not exist, create new .ini file. Path: " + sIniFilePath);
                tpc.F_WritePrivateProfileString("ProjectName", "Ground PC or Primary PC", ProjectName.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("ProjectName", "Cloud PC or Backup PC", "CTestProject", sIniFilePath);
                tpc.F_WritePrivateProfileString("IP", "Ground PC or Primary PC", WebAccessIP.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("IP", "Cloud PC or Backup PC", "172.18.3.65", sIniFilePath);
            }
        }

    }
}
