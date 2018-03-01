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
using System.Collections.ObjectModel;

namespace CreateEventLogData
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
        string sTestItemName = "CreateEventLogData";
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

            //Create Data Log Trend
            if (bPartResult == true)
            {
                
                sw.Reset(); sw.Start();
                try
                {
                    CheckData();
                    
                    if (bPartResult == true)
                        CreateEventLogData(sPrimaryProject, sLanguage);
                    if (bPartResult == true)
                        CreateEventLogData2(sPrimaryProject, sLanguage);  // 記錄1,3,5,7,9 中間有空隔不連續 ( AE feedback)
                    if (bPartResult == true)
                        CreateEventLogData3(sPrimaryProject, sLanguage);  // 只記錄1個測點 (AE feedback)
                    
                }
                catch (Exception ex)
                {
                    EventLog.AddLog(@"Error occurred create EventLogData: " + ex.ToString());
                    bPartResult = false;
                }
                sw.Stop();
                PrintStep("Create", "Create EventLogData", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());

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

        /*
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
        */

        private void CreateEventLogData(string sProjectName, string slanguage)
        {
            try
            {
                EventLog.AddLog("Create EventLogData");
                driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/eventLog/eveLogPg.asp') and contains(@href, 'pos=add')]")).Click();

                switch (slanguage)
                {
                    case "ENG":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("Event Tag == Reference Value");
                        break;
                    case "CHT":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("事件測點 == 參考值");
                        break;
                    case "CHS":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("事件点 == 参考值");
                        break;
                    case "JPN":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("ｲﾍﾞﾝﾄﾀｸﾞ == 参照値");
                        break;
                    case "KRN":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("이벤트 태그 == Reference Value");
                        break;
                    case "FRN":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("Repère d'Evénement == Valeur de Référence");
                        break;

                    default:
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("Event Tag == Reference Value");
                        break;
                }

                driver.FindElement(By.Name("EventLogName")).Clear();
                driver.FindElement(By.Name("EventLogName")).SendKeys("EventLog_" + sProjectName);
                driver.FindElement(By.Name("EventTag")).Clear();
                driver.FindElement(By.Name("EventTag")).SendKeys("ConAna_0241");
                driver.FindElement(By.Name("EventRefVal")).Clear();
                driver.FindElement(By.Name("EventRefVal")).SendKeys("51");
                EventLog.AddLog("Set Event Log trigger event");

                driver.FindElement(By.Name("EventLogTag")).Click();

                new SelectElement(driver.FindElement(By.Name("ChSel"))).SelectByText("240");
                Thread.Sleep(2000);
                driver.SwitchTo().Alert().Accept();
                Thread.Sleep(1000);

                for (int i = 1; i <= 9; i++)
                {
                    try
                    {
                        driver.FindElement(By.Name(string.Format("TagName{0}", i))).SendKeys(string.Format("ConAna_000{0}", i));
                    }
                    catch (Exception ex)
                    {
                        EventLog.AddLog("CreateEventLogData 1~9 error: " + ex.ToString());
                    }
                }
                EventLog.AddLog("Create 1~9  tags to log");

                for (int i = 10; i <= 99; i++)
                {
                    try
                    {
                        driver.FindElement(By.Name(string.Format("TagName{0}", i))).SendKeys(string.Format("ConAna_00{0}", i));
                    }
                    catch (Exception ex)
                    {
                        EventLog.AddLog("CreateEventLogData 10~99 error: " + ex.ToString());
                    }
                }
                EventLog.AddLog("Create 10~99  tags to log");

                for (int i = 100; i <= 240; i++)
                {
                    try
                    {
                        if (i == 240)
                        {
                            driver.FindElement(By.Name(string.Format("TagName{0}", i))).SendKeys(string.Format("ConAna_0{0}", i));
                            driver.FindElement(By.Name("submit")).Click();
                        }
                        else
                        {
                            driver.FindElement(By.Name(string.Format("TagName{0}", i))).SendKeys(string.Format("ConAna_0{0}", i));
                        }
                    }
                    catch (Exception ex)
                    {
                        EventLog.AddLog("CreateEventLogData 100~240 error: " + ex.ToString());
                    }
                }
                EventLog.AddLog("Create 100~240  tags to log");

                Thread.Sleep(1000);

                driver.FindElement(By.Name("EventRefVal")).SendKeys("");
                driver.FindElement(By.Name("submit")).Click();

                Thread.Sleep(1000);
            }
            catch (Exception ex)
            {
                EventLog.AddLog("CreateEventLogData error: " + ex.ToString());
                bPartResult = false;
            }
        }

        private void CreateEventLogData2(string sProjectName, string slanguage)
        {
            try
            {
                EventLog.AddLog("Create EventLogData2");
                driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/eventLog/eveLogPg.asp') and contains(@href, 'pos=add')]")).Click();

                switch (slanguage)
                {
                    case "ENG":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("Event Tag == Reference Value");
                        break;
                    case "CHT":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("事件測點 == 參考值");
                        break;
                    case "CHS":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("事件点 == 参考值");
                        break;
                    case "JPN":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("ｲﾍﾞﾝﾄﾀｸﾞ == 参照値");
                        break;
                    case "KRN":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("이벤트 태그 == Reference Value");
                        break;
                    case "FRN":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("Repère d'Evénement == Valeur de Référence");
                        break;

                    default:
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("Event Tag == Reference Value");
                        break;

                }

                driver.FindElement(By.Name("EventLogName")).Clear();
                driver.FindElement(By.Name("EventLogName")).SendKeys("EventLog13579_" + sProjectName);
                driver.FindElement(By.Name("EventTag")).Clear();
                driver.FindElement(By.Name("EventTag")).SendKeys("ConAna_0241");
                driver.FindElement(By.Name("EventRefVal")).Clear();
                driver.FindElement(By.Name("EventRefVal")).SendKeys("51");
                EventLog.AddLog("Set Event Log trigger event");

                driver.FindElement(By.Name("EventLogTag")).Click();

                Thread.Sleep(1000);

                for (int i = 1; i <= 9; i = i + 2)
                {
                    try
                    {
                        if (i == 9)
                        {
                            driver.FindElement(By.Name(string.Format("TagName{0}", i))).SendKeys(string.Format("ConAna_000{0}", i));
                            driver.FindElement(By.Name("submit")).Click();
                        }
                        else
                            driver.FindElement(By.Name(string.Format("TagName{0}", i))).SendKeys(string.Format("ConAna_000{0}", i));
                    }
                    catch (Exception ex)
                    {
                        EventLog.AddLog("CreateEventLogData 1,3,5,7,9 error: " + ex.ToString());
                    }
                }
                EventLog.AddLog("Create 1,3,5,7,9  tags to log");

                Thread.Sleep(1000);

                driver.FindElement(By.Name("EventRefVal")).SendKeys("");
                driver.FindElement(By.Name("submit")).Click();

                Thread.Sleep(1000);
            }
            catch (Exception ex)
            {
                EventLog.AddLog("CreateEventLogData2 error: " + ex.ToString());
                bPartResult = false;
            }
        }

        private void CreateEventLogData3(string sProjectName, string slanguage)
        {
            try
            {
                EventLog.AddLog("Create EventLogData3");
                driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/eventLog/eveLogPg.asp') and contains(@href, 'pos=add')]")).Click();

                switch (slanguage)
                {
                    case "ENG":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("Event Tag == Reference Value");
                        break;
                    case "CHT":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("事件測點 == 參考值");
                        break;
                    case "CHS":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("事件点 == 参考值");
                        break;
                    case "JPN":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("ｲﾍﾞﾝﾄﾀｸﾞ == 参照値");
                        break;
                    case "KRN":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("이벤트 태그 == Reference Value");
                        break;
                    case "FRN":
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("Repère d'Evénement == Valeur de Référence");
                        break;

                    default:
                        new SelectElement(driver.FindElement(By.Name("EventTypeSel"))).SelectByText("Event Tag == Reference Value");
                        break;

                }

                driver.FindElement(By.Name("EventLogName")).Clear();
                driver.FindElement(By.Name("EventLogName")).SendKeys("EventLog1_" + sProjectName);
                driver.FindElement(By.Name("EventTag")).Clear();
                driver.FindElement(By.Name("EventTag")).SendKeys("ConAna_0241");
                driver.FindElement(By.Name("EventRefVal")).Clear();
                driver.FindElement(By.Name("EventRefVal")).SendKeys("51");
                EventLog.AddLog("Set Event Log trigger event");

                driver.FindElement(By.Name("EventLogTag")).Click();

                Thread.Sleep(1000);

                try
                {
                    driver.FindElement(By.Name(string.Format("TagName1"))).SendKeys(string.Format("ConAna_0001"));
                    driver.FindElement(By.Name("submit")).Click();
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("CreateEventLogData 1,3,5,7,9 error: " + ex.ToString());
                }

                EventLog.AddLog("Create 1 tags to log");

                Thread.Sleep(1000);

                driver.FindElement(By.Name("EventRefVal")).SendKeys("");
                driver.FindElement(By.Name("submit")).Click();

                Thread.Sleep(1000);
            }
            catch (Exception ex)
            {
                EventLog.AddLog("CreateEventLogData3 error: " + ex.ToString());
                bPartResult = false;
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

        private void CheckData()
        {
            try
            {
                EventLog.AddLog("Check Data");
                driver.SwitchTo().Frame("rightFrame");

                driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/eventlog/EveLogList.asp')]")).Click();
                Thread.Sleep(3000);

                List<string> matchingLinks = new List<string>();
                ReadOnlyCollection<IWebElement> links = driver.FindElements(By.XPath("//a[contains(@href, 'EveLogDel.asp?nid')]"));

                for(int i= 0;i<links.Count;i++)
                {
                    driver.FindElement(By.XPath("/html/body/font/table/tbody/tr[3]/td/center/table/tbody/tr[2]/td[4]/font/a")).Click();
                    Thread.Sleep(1000);
                    driver.SwitchTo().Alert().Accept();
                    Thread.Sleep(1000);
                }
            }
            catch (Exception ex)
            {
                EventLog.AddLog("CheckData error: " + ex.ToString());
                bPartResult = false;
            }
        }
    }
}