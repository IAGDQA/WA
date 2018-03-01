using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using ThirdPartyToolControl;
using iATester;
using CommonFunction;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;       // for SelectElement use

namespace View_and_Save_DataLogTrendData
{
    public partial class Form1 : Form, iATester.iCom
    {
        cThirdPartyToolControl tpc = new cThirdPartyToolControl();
        cWACommonFunction wcf = new cWACommonFunction();
        cEventLog EventLog = new cEventLog();
        Stopwatch sw = new Stopwatch();

        private IWebDriver driver;
        int iRetryNum;
        bool bFinalResult = true;
        bool bPartResult = true;
        string baseUrl;
        string sTestItemName = "View_and_Save_DataLogTrendData";
        string sIniFilePath = @"C:\WebAccessAutoTestSettingInfo.ini";
        string sTestLogFolder = @"C:\WALogData";

        //Send Log data to iAtester
        public event EventHandler<LogEventArgs> eLog = delegate { };
        //Send test result to iAtester
        public event EventHandler<ResultEventArgs> eResult = delegate { };
        //Send execution status to iAtester
        public event EventHandler<StatusEventArgs> eStatus = delegate { };

        public void StartTest()
        {
            //Add test code
            long lErrorCode = 0;
            EventLog.AddLog(string.Format("***** {0} test start (by iATester) *****", sTestItemName));
            CheckifIniFileChange();
            EventLog.AddLog("Primary Project= " + textBox_Primary_project.Text);
            EventLog.AddLog("Primary IP= " + textBox_Primary_IP.Text);
            EventLog.AddLog("Secondary Project= " + textBox_Secondary_project.Text);
            EventLog.AddLog("Secondary IP= " + textBox_Secondary_IP.Text);
            //Form1_Load(textBox_Primary_project.Text, textBox_Primary_IP.Text, textBox_Secondary_project.Text, textBox_Secondary_IP.Text, sTestLogFolder, comboBox_Browser.Text, textbox_UserEmail.Text, comboBox_Language.Text);
            for (int i = 0; i < iRetryNum; i++)
            {
                EventLog.AddLog(string.Format("===Retry Number : {0} / {1} ===", i + 1, iRetryNum));
                lErrorCode = Form1_Load(textBox_Primary_project.Text, textBox_Primary_IP.Text, textBox_Secondary_project.Text, textBox_Secondary_IP.Text, sTestLogFolder, comboBox_Browser.Text, textbox_UserEmail.Text, comboBox_Language.Text);
                if (lErrorCode == 0)
                {
                    eResult(this, new ResultEventArgs(iResult.Pass));
                    break;
                }
                else
                {
                    if (i == iRetryNum - 1)
                        eResult(this, new ResultEventArgs(iResult.Fail));
                }
            }

            eStatus(this, new StatusEventArgs(iStatus.Completion));

            EventLog.AddLog(string.Format("***** {0} test end (by iATester) *****", sTestItemName));
        }

        private void Start_Click(object sender, EventArgs e)
        {
            EventLog.AddLog(string.Format("***** {0} test start *****", sTestItemName));
            CheckifIniFileChange();
            EventLog.AddLog("Primary Project= " + textBox_Primary_project.Text);
            EventLog.AddLog("Primary IP= " + textBox_Primary_IP.Text);
            EventLog.AddLog("Secondary Project= " + textBox_Secondary_project.Text);
            EventLog.AddLog("Secondary IP= " + textBox_Secondary_IP.Text);
            Form1_Load(textBox_Primary_project.Text, textBox_Primary_IP.Text, textBox_Secondary_project.Text, textBox_Secondary_IP.Text, sTestLogFolder, comboBox_Browser.Text, textbox_UserEmail.Text, comboBox_Language.Text);
            EventLog.AddLog(string.Format("***** {0} test end *****", sTestItemName));
        }

        public Form1()
        {
            InitializeComponent();

            comboBox_Browser.SelectedIndex = 0;
            comboBox_Language.SelectedIndex = 0;
            Text = string.Format("Advantech WebAccess Auto Test ( {0} )", sTestItemName);
            if (System.IO.File.Exists(sIniFilePath))
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
        }

        long Form1_Load(string sPrimaryProject, string sPrimaryIP, string sSecondaryProject, string sSecondaryIP, string sTestLogFolder, string sBrowser, string sUserEmail, string sLanguage)
        {
            bPartResult = true;
            baseUrl = "http://" + sPrimaryIP;
            if (bPartResult == true)
            {
                EventLog.AddLog("Open browser for selenium driver use");
                sw.Reset(); sw.Start();
                try
                {
                    if (sBrowser == "Internet Explorer")
                    {
                        EventLog.AddLog("Browser= Internet Explorer");
                        InternetExplorerOptions options = new InternetExplorerOptions();
                        options.IgnoreZoomLevel = true;
                        driver = new InternetExplorerDriver(options);
                        driver.Manage().Window.Maximize();
                    }
                    else
                    {
                        EventLog.AddLog("Not support temporary");
                        bPartResult = false;
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog(@"Error opening browser: " + ex.ToString());
                    bPartResult = false;
                }
                sw.Stop();
                PrintStep("Open browser", "Open browser for selenium driver use", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());
            }

            //Login test
            if (bPartResult == true)
            {
                EventLog.AddLog("Login WebAccess homepage");
                sw.Reset(); sw.Start();
                try
                {
                    driver.Navigate().GoToUrl(baseUrl + "/broadWeb/bwRoot.asp?username=admin");
                    driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwconfig.asp?username=admin')]")).Click();
                    driver.FindElement(By.Id("userField")).Submit();
                    Thread.Sleep(3000);
                    driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sPrimaryProject + "')]")).Click();
                }
                catch (Exception ex)
                {
                    EventLog.AddLog(@"Error occurred logging on: " + ex.ToString());
                    bPartResult = false;
                }
                sw.Stop();
                PrintStep("Login", "Login project manager page", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());

                Thread.Sleep(1000);
            }

            //View data log Trend
            if (bPartResult == true)
            {
                EventLog.AddLog("View data log Trend");
                sw.Reset(); sw.Start();
                try
                {
                    EventLog.AddLog("Start view data log trend");
                    StartViewDataLogTrend(sPrimaryProject, sTestLogFolder, sLanguage);
                }
                catch (Exception ex)
                {
                    EventLog.AddLog(@"Error occurred View data log Trend: " + ex.ToString());
                    bPartResult = false;
                }
                sw.Stop();
                PrintStep("View", "View data log Trend", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());

                Thread.Sleep(1000);
            }

            driver.Dispose();

            #region Result judgement
            if (bFinalResult && bPartResult)
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

        private void StartViewDataLogTrend(string sProjectName, string sTestLogFolder, string slanguage)
        {
            // Start view
            EventLog.AddLog("Start view");
            driver.SwitchTo().Frame("rightFrame");
            //driver.FindElement(By.XPath("//tr[2]/td/a/font")).Click();
            driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/system/bwview.asp')]")).Click();

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
            {
                EventLog.AddLog("Cannot get Start View WebAccess Main Page handle");
                bPartResult = false;
            }

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
            {
                EventLog.AddLog("Cannot get Login keyboard handle");
                bPartResult = false;
            }

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
            {
                EventLog.AddLog("Cannot get DataLog Trend List handle");
                bPartResult = false;
            }
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

                EventLog.PrintScreen(string.Format("DataLogTrend_Interval_{0}_RightMost_Step1", sInterval[iInterval - 1]));

                SendKeys.SendWait("{F2}"); // Left
                Thread.Sleep(4000);

                EventLog.PrintScreen(string.Format("DataLogTrend_Interval_{0}_Left_Step2", sInterval[iInterval - 1]));

                SendKeys.SendWait("{F2}"); // Left
                Thread.Sleep(4000);

                EventLog.PrintScreen(string.Format("DataLogTrend_Interval_{0}_Right_Step3", sInterval[iInterval - 1]));

                SendKeys.SendWait("{F3}"); // Right
                Thread.Sleep(4000);

                EventLog.PrintScreen(string.Format("DataLogTrend_Interval_{0}_Right_Step4", sInterval[iInterval - 1]));

                SendKeys.SendWait("{F3}"); // Right
                Thread.Sleep(4000);

                EventLog.PrintScreen(string.Format("DataLogTrend_Interval_{0}_Right_Step5", sInterval[iInterval - 1]));

                SendKeys.SendWait("{F1}"); // Left most
                Thread.Sleep(4000);

                EventLog.PrintScreen(string.Format("DataLogTrend_Interval_{0}_LeftMost_Step6", sInterval[iInterval - 1]));

                SendKeys.SendWait("{F4}"); // Right most
                Thread.Sleep(4000);

                EventLog.PrintScreen(string.Format("DataLogTrend_Interval_{0}_RightMost_Step7", sInterval[iInterval - 1]));

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
                {
                    EventLog.AddLog("Cannot get EnterText_PointInfo handle");
                    bPartResult = false;
                }


                int iChange_Button_of_PointInfo = 0;
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
                {
                    EventLog.AddLog("Cannot get iChange_Button_of_PointInfo handle");
                    bPartResult = false;
                }

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
                {
                    EventLog.AddLog("Cannot get iEnterValue_of_interval handle");
                    bPartResult = false;
                }


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
                {
                    EventLog.AddLog("Cannot get iEnterButton_ChangeWindow handle");
                    bPartResult = false;
                }

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
                {
                    EventLog.AddLog("Cannot get iExitButton_PointInfo handle");
                    bPartResult = false;
                }

                if (iInterval == sInterval.Length)
                {
                    Thread.Sleep(4000);
                    SendKeys.SendWait("{F4}"); // Right most
                    Thread.Sleep(4000);

                    EventLog.PrintScreen(string.Format("DataLogTrend_Interval_{0}_RightMost_Step1", sInterval[iInterval - 1]));

                    SendKeys.SendWait("{F2}"); // Left
                    Thread.Sleep(4000);

                    EventLog.PrintScreen(string.Format("DataLogTrend_Interval_{0}_Left_Step2", sInterval[iInterval - 1]));

                    SendKeys.SendWait("{F2}"); // Left
                    Thread.Sleep(4000);

                    EventLog.PrintScreen(string.Format("DataLogTrend_Interval_{0}_Right_Step3", sInterval[iInterval - 1]));

                    SendKeys.SendWait("{F3}"); // Right
                    Thread.Sleep(4000);

                    EventLog.PrintScreen(string.Format("DataLogTrend_Interval_{0}_Right_Step4", sInterval[iInterval - 1]));

                    SendKeys.SendWait("{F3}"); // Right
                    Thread.Sleep(4000);

                    EventLog.PrintScreen(string.Format("DataLogTrend_Interval_{0}_Right_Step5", sInterval[iInterval - 1]));

                    SendKeys.SendWait("{F1}"); // Left most
                    Thread.Sleep(4000);

                    EventLog.PrintScreen(string.Format("DataLogTrend_Interval_{0}_LeftMost_Step6", sInterval[iInterval - 1]));

                    SendKeys.SendWait("{F4}"); // Right most
                    Thread.Sleep(4000);

                    EventLog.PrintScreen(string.Format("DataLogTrend_Interval_{0}_RightMost_Step7", sInterval[iInterval - 1]));
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
            {
                EventLog.AddLog("Cannot get Start View WebAccess Main Page handle");
                bPartResult = false;
            }

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
                {
                    EventLog.AddLog("Cannot get $view001.htm file");
                    bPartResult = false;
                }
            }

            if (iWA_MainPage > 0)
            {
                Thread.Sleep(5000); //delay 5s before close $view001 window
                EventLog.AddLog("Send alt+F4 keyboard command to close $view001.htm window");
                SendKeys.SendWait("%{F4}");
                Thread.Sleep(500);
            }
            else
            {
                EventLog.AddLog("Cannot find $view001.htm window");
                bPartResult = false;
            }

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

        private void PrintStep(string sTestItem, string sDescription, bool bResult, string sErrorCode, string sExTime)
        {
            EventLog.AddLog(string.Format("UI Result: {0},{1},{2},{3},{4}", sTestItem, sDescription, bResult, sErrorCode, sExTime));
        }

        private void InitialRequiredInfo(string sFilePath)
        {
            StringBuilder sDefaultUserLanguage = new StringBuilder(255);
            StringBuilder sDefaultUserEmail = new StringBuilder(255);
            StringBuilder sDefaultUserRetryNum = new StringBuilder(255);
            StringBuilder sBrowser = new StringBuilder(255);
            StringBuilder sDefaultProjectName1 = new StringBuilder(255);
            StringBuilder sDefaultProjectName2 = new StringBuilder(255);
            StringBuilder sDefaultIP1 = new StringBuilder(255);
            StringBuilder sDefaultIP2 = new StringBuilder(255);

            tpc.F_GetPrivateProfileString("UserInfo", "Language", "NA", sDefaultUserLanguage, 255, sFilePath);
            tpc.F_GetPrivateProfileString("UserInfo", "Email", "NA", sDefaultUserEmail, 255, sFilePath);
            tpc.F_GetPrivateProfileString("UserInfo", "RetryNum", "NA", sDefaultUserRetryNum, 255, sFilePath);
            tpc.F_GetPrivateProfileString("UserInfo", "Browser", "NA", sBrowser, 255, sFilePath);
            tpc.F_GetPrivateProfileString("ProjectName", "Primary PC", "NA", sDefaultProjectName1, 255, sFilePath);
            tpc.F_GetPrivateProfileString("ProjectName", "Secondary PC", "NA", sDefaultProjectName2, 255, sFilePath);
            tpc.F_GetPrivateProfileString("IP", "Primary PC", "NA", sDefaultIP1, 255, sFilePath);
            tpc.F_GetPrivateProfileString("IP", "Secondary PC", "NA", sDefaultIP2, 255, sFilePath);

            comboBox_Language.Text = sDefaultUserLanguage.ToString();
            textbox_UserEmail.Text = sDefaultUserEmail.ToString();
            comboBox_Browser.Text = sBrowser.ToString();
            textBox_Primary_project.Text = sDefaultProjectName1.ToString();
            textBox_Secondary_project.Text = sDefaultProjectName2.ToString();
            textBox_Primary_IP.Text = sDefaultIP1.ToString();
            textBox_Secondary_IP.Text = sDefaultIP2.ToString();
            if (Int32.TryParse(sDefaultUserRetryNum.ToString(), out iRetryNum))     // 在這邊取得retry number
            {
                EventLog.AddLog("Converted retry number '{0}' to {1}.", sDefaultUserRetryNum.ToString(), iRetryNum);
            }
            else
            {
                EventLog.AddLog("Attempted conversion of '{0}' failed.",
                                sDefaultUserRetryNum.ToString() == null ? "<null>" : sDefaultUserRetryNum.ToString());
                EventLog.AddLog("Set the number of retry as 3");
                iRetryNum = 3;  // 轉換失敗 直接指定預設值為3
            }
        }

        private void CheckifIniFileChange()
        {
            StringBuilder sDefaultUserLanguage = new StringBuilder(255);
            StringBuilder sDefaultUserEmail = new StringBuilder(255);
            StringBuilder sDefaultUserRetryNum = new StringBuilder(255);
            StringBuilder sBrowser = new StringBuilder(255);
            StringBuilder sDefaultProjectName1 = new StringBuilder(255);
            StringBuilder sDefaultProjectName2 = new StringBuilder(255);
            StringBuilder sDefaultIP1 = new StringBuilder(255);
            StringBuilder sDefaultIP2 = new StringBuilder(255);

            if (System.IO.File.Exists(sIniFilePath))    // 比對ini檔與ui上的值是否相同
            {
                EventLog.AddLog(".ini file exist, check if .ini file need to update");
                tpc.F_GetPrivateProfileString("UserInfo", "Language", "NA", sDefaultUserLanguage, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("UserInfo", "Email", "NA", sDefaultUserEmail, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("UserInfo", "RetryNum", "NA", sDefaultUserRetryNum, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("UserInfo", "Browser", "NA", sBrowser, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("ProjectName", "Primary PC", "NA", sDefaultProjectName1, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("ProjectName", "Secondary PC", "NA", sDefaultProjectName2, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("IP", "Primary PC", "NA", sDefaultIP1, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("IP", "Secondary PC", "NA", sDefaultIP2, 255, sIniFilePath);

                if (comboBox_Language.Text != sDefaultUserLanguage.ToString())
                {
                    tpc.F_WritePrivateProfileString("UserInfo", "Language", comboBox_Language.Text, sIniFilePath);
                    EventLog.AddLog("New Language update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultUserLanguage.ToString());
                    EventLog.AddLog("New ini:" + comboBox_Language.Text);
                }
                if (textbox_UserEmail.Text != sDefaultUserEmail.ToString())
                {
                    tpc.F_WritePrivateProfileString("UserInfo", "Email", textbox_UserEmail.Text, sIniFilePath);
                    EventLog.AddLog("New UserEmail update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultUserEmail.ToString());
                    EventLog.AddLog("New ini:" + textbox_UserEmail.Text);
                }
                if (comboBox_Browser.Text != sBrowser.ToString())
                {
                    tpc.F_WritePrivateProfileString("UserInfo", "Browser", comboBox_Browser.Text, sIniFilePath);
                    EventLog.AddLog("New Browser update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sBrowser.ToString());
                    EventLog.AddLog("New ini:" + comboBox_Browser.Text);
                }
                if (textBox_Primary_project.Text != sDefaultProjectName1.ToString())
                {
                    tpc.F_WritePrivateProfileString("ProjectName", "Primary PC", textBox_Primary_project.Text, sIniFilePath);
                    EventLog.AddLog("New Primary ProjectName update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultProjectName1.ToString());
                    EventLog.AddLog("New ini:" + textBox_Primary_project.Text);
                }
                if (textBox_Secondary_project.Text != sDefaultProjectName2.ToString())
                {
                    tpc.F_WritePrivateProfileString("ProjectName", "Secondary PC", textBox_Secondary_project.Text, sIniFilePath);
                    EventLog.AddLog("New Secondary ProjectName update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultProjectName2.ToString());
                    EventLog.AddLog("New ini:" + textBox_Secondary_project.Text);
                }
                if (textBox_Primary_IP.Text != sDefaultIP1.ToString())
                {
                    tpc.F_WritePrivateProfileString("IP", "Primary PC", textBox_Primary_IP.Text, sIniFilePath);
                    EventLog.AddLog("New Primary IP update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultIP1.ToString());
                    EventLog.AddLog("New ini:" + textBox_Primary_IP.Text);
                }
                if (textBox_Secondary_IP.Text != sDefaultIP2.ToString())
                {
                    tpc.F_WritePrivateProfileString("IP", "Secondary PC", textBox_Secondary_IP.Text, sIniFilePath);
                    EventLog.AddLog("New Secondary IP update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultIP2.ToString());
                    EventLog.AddLog("New ini:" + textBox_Secondary_IP.Text);
                }
            }
            else
            {   // 若ini檔不存在 則建立新的
                EventLog.AddLog(".ini file not exist, create new .ini file. Path: " + sIniFilePath);
                tpc.F_WritePrivateProfileString("UserInfo", "Language", comboBox_Language.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("UserInfo", "Email", textbox_UserEmail.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("UserInfo", "RetryNum", "3", sIniFilePath);
                tpc.F_WritePrivateProfileString("UserInfo", "Browser", comboBox_Browser.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("ProjectName", "Primary PC", textBox_Primary_project.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("ProjectName", "Secondary PC", textBox_Secondary_project.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("IP", "Primary PC", textBox_Primary_IP.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("IP", "Secondary PC", textBox_Secondary_IP.Text, sIniFilePath);
            }
        }
    }
}
