using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.IO;
using System.Diagnostics;
using ThirdPartyToolControl;
using iATester;
using CommonFunction;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;       // for SelectElement use

namespace CreateGlobalScriptData
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
        string sTestItemName = "CreateGlobalScriptData";
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

            //Step 0: Download
            if (bPartResult == true)
            {
                try
                {
                    EventLog.AddLog("Download...");
                    bPartResult = wcf.Download(driver, sLanguage);
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Download error: " + ex.ToString());
                    bPartResult = false;
                }
                Thread.Sleep(1000);
            }

            //Step1: Copy "ConstTag_Set.scr" and "alm_set_ConAna_51.scr" and "alm_ack.scr" 
            //        to C:\WebAccess\Node\config\ProjectName\bgr  and  C:\WebAccess\Node\ProjectName\bgr
            if (bPartResult == true)
            {
                try
                {
                    EventLog.AddLog("Copy files...");
                    //string sCurrentFilePath = Directory.GetCurrentDirectory();
                    string sCurrentFilePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetAssembly(this.GetType()).Location);

                    string sourceFile1 = sCurrentFilePath + "\\GlobalScriptSample\\ConstTag_Set.scr";
                    string destFile1_1 = string.Format("C:\\WebAccess\\Node\\config\\{0}_TestSCADA\\bgr\\ConstTag_Set.scr", sPrimaryProject);
                    string destFile1_2 = string.Format("C:\\WebAccess\\Node\\{0}_TestSCADA\\bgr\\ConstTag_Set.scr", sPrimaryProject);

                    string sourceFile2 = sCurrentFilePath + "\\GlobalScriptSample\\alm_set_ConAna_51.scr";
                    string destFile2_1 = string.Format("C:\\WebAccess\\Node\\config\\{0}_TestSCADA\\bgr\\alm_set_ConAna_51.scr", sPrimaryProject);
                    string destFile2_2 = string.Format("C:\\WebAccess\\Node\\{0}_TestSCADA\\bgr\\alm_set_ConAna_51.scr", sPrimaryProject);

                    string sourceFile3 = sCurrentFilePath + "\\GlobalScriptSample\\alm_ack.scr";
                    string destFile3_1 = string.Format("C:\\WebAccess\\Node\\config\\{0}_TestSCADA\\bgr\\alm_ack.scr", sPrimaryProject);
                    string destFile3_2 = string.Format("C:\\WebAccess\\Node\\{0}_TestSCADA\\bgr\\alm_ack.scr", sPrimaryProject);

                    string sourceFile4 = sCurrentFilePath + "\\GlobalScriptSample\\ConstTxt_Set.scr";
                    string destFile4_1 = string.Format("C:\\WebAccess\\Node\\config\\{0}_TestSCADA\\bgr\\ConstTxt_Set.scr", sPrimaryProject);
                    string destFile4_2 = string.Format("C:\\WebAccess\\Node\\{0}_TestSCADA\\bgr\\ConstTxt_Set.scr", sPrimaryProject);

                    System.IO.File.Copy(sourceFile1, destFile1_1, true);
                    System.IO.File.Copy(sourceFile1, destFile1_2, true);
                    System.IO.File.Copy(sourceFile2, destFile2_1, true);
                    System.IO.File.Copy(sourceFile2, destFile2_2, true);
                    System.IO.File.Copy(sourceFile3, destFile3_1, true);
                    System.IO.File.Copy(sourceFile3, destFile3_2, true);
                    System.IO.File.Copy(sourceFile4, destFile4_1, true);
                    System.IO.File.Copy(sourceFile4, destFile4_2, true);
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Copy files error: " + ex.ToString());
                    bPartResult = false;
                }
                Thread.Sleep(1000);
            }

            //Step2: Set global script
            if (bPartResult == true)
            {
                try
                {
                    EventLog.AddLog("Set global script...");
                    CreateGlobalScript(sLanguage);
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Set global script error: " + ex.ToString());
                    bPartResult = false;
                }
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

        private void CreateGlobalScript(string sLanguage)
        {
            try
            {
                driver.SwitchTo().Frame("rightFrame");
                driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/GbScript/GbScriptPg.asp')]")).Click();

                switch (sLanguage)
                {
                    case "ENG":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_1"))).SelectByText("Enable");
                        break;
                    case "CHT":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_1"))).SelectByText("啟用");
                        break;
                    case "CHS":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_1"))).SelectByText("允许");
                        break;
                    case "JPN":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_1"))).SelectByText("稼動");
                        break;
                    case "KRN":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_1"))).SelectByText("가동");
                        break;
                    case "FRN":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_1"))).SelectByText("Activer");
                        break;

                    default:
                        new SelectElement(driver.FindElement(By.Name("StatusSel_1"))).SelectByText("Enable");
                        break;
                }
                driver.FindElement(By.Name("Description_1")).Clear();
                driver.FindElement(By.Name("Description_1")).SendKeys("Set each Const tag plus 1");
                driver.FindElement(By.Name("RunScript_1")).Clear();
                driver.FindElement(By.Name("RunScript_1")).SendKeys("ConstTag_Set.scr");
                driver.FindElement(By.Name("RunInterval_1")).Clear();
                driver.FindElement(By.Name("RunInterval_1")).SendKeys("200");

                switch (sLanguage)
                {
                    case "ENG":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_2"))).SelectByText("Enable");
                        break;
                    case "CHT":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_2"))).SelectByText("啟用");
                        break;
                    case "CHS":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_2"))).SelectByText("允许");
                        break;
                    case "JPN":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_2"))).SelectByText("稼動");
                        break;
                    case "KRN":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_2"))).SelectByText("가동");
                        break;
                    case "FRN":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_2"))).SelectByText("Activer");
                        break;

                    default:
                        new SelectElement(driver.FindElement(By.Name("StatusSel_2"))).SelectByText("Enable");
                        break;
                }
                driver.FindElement(By.Name("Description_2")).Clear();
                driver.FindElement(By.Name("Description_2")).SendKeys("Set all ConAna 51");
                driver.FindElement(By.Name("StartScript_2")).Clear();
                driver.FindElement(By.Name("StartScript_2")).SendKeys("alm_set_ConAna_51.scr");

                switch (sLanguage)
                {
                    case "ENG":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_3"))).SelectByText("Enable");
                        break;
                    case "CHT":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_3"))).SelectByText("啟用");
                        break;
                    case "CHS":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_3"))).SelectByText("允许");
                        break;
                    case "JPN":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_3"))).SelectByText("稼動");
                        break;
                    case "KRN":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_3"))).SelectByText("가동");
                        break;
                    case "FRN":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_3"))).SelectByText("Activer");
                        break;

                    default:
                        new SelectElement(driver.FindElement(By.Name("StatusSel_3"))).SelectByText("Enable");
                        break;
                }
                driver.FindElement(By.Name("RunScript_3")).Clear();
                driver.FindElement(By.Name("RunScript_3")).SendKeys("alm_ack.scr");
                driver.FindElement(By.Name("RunInterval_3")).Clear();
                driver.FindElement(By.Name("RunInterval_3")).SendKeys("2400");   // 2400 = 60s

                switch (sLanguage)
                {
                    case "ENG":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_4"))).SelectByText("Enable");
                        break;
                    case "CHT":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_4"))).SelectByText("啟用");
                        break;
                    case "CHS":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_4"))).SelectByText("允许");
                        break;
                    case "JPN":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_4"))).SelectByText("稼動");
                        break;
                    case "KRN":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_4"))).SelectByText("가동");
                        break;
                    case "FRN":
                        new SelectElement(driver.FindElement(By.Name("StatusSel_4"))).SelectByText("Activer");
                        break;

                    default:
                        new SelectElement(driver.FindElement(By.Name("StatusSel_4"))).SelectByText("Enable");
                        break;
                }
                driver.FindElement(By.Name("StopScript_4")).Clear();
                driver.FindElement(By.Name("StopScript_4")).SendKeys("ConstTxt_Set.scr");

                driver.FindElement(By.Name("Submit")).Click();
            }
            catch (Exception ex)
            {
                EventLog.AddLog("CreateGlobalScript error: " + ex.ToString());
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
    }
}
