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

namespace CreateProjectSCADANode
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
        string sTestItemName = "CreateProjectSCADANode";
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
                    Thread.Sleep(2000);
                    //driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sProjectName + "')]")).Click();
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
            
            // Step 0: Set LogData Maintenance
            if (bPartResult == true)
            {
                EventLog.AddLog("Set LogData Maintenance");
                sw.Reset(); sw.Start();
                try
                {
                    SetLogDataMaintenance(sTestLogFolder);
                    Thread.Sleep(1500);
                }
                catch (Exception ex)
                {
                    EventLog.AddLog(@"Error occurred when set LogData Maintenance: " + ex.ToString());
                    bPartResult = false;
                }
                sw.Stop();
                PrintStep("Set LogData Maintenance", "Set LogData Maintenance", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());
            }

            // Step 1: Create project and scada
            if (bPartResult == true)
            {
                EventLog.AddLog("Create project and scada");
                sw.Reset(); sw.Start();
                try
                {
                    if (wcf.IsTestElementPresent(driver, "XPath", "//a[contains(@href, '/broadWeb/project/deleteProject.asp?') and contains(@href, 'ProjName=" + sPrimaryProject + "')]"))
                    {
                        EventLog.AddLog("Delete old project of same name");
                        DeleteProject(sPrimaryProject, sLanguage);
                    }

                    long lStartTime = Environment.TickCount;
                    long lEndTime = 0;
                    do
                    {
                        if (!wcf.IsTestElementPresent(driver, "XPath", "//a[contains(@href, '/broadWeb/project/deleteProject.asp?') and contains(@href, 'ProjName=" + sPrimaryProject + "')]"))
                            break;
                        lEndTime = Environment.TickCount;
                    } while ((lEndTime - lStartTime) <= 60000);

                    if (lEndTime - lStartTime <= 60000)
                    {
                        CreateProject(sPrimaryProject, sPrimaryIP, sLanguage);
                        Thread.Sleep(3000);
                        driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sPrimaryProject + "')]")).Click();
                        Thread.Sleep(1000);
                        CreateSCADANode(sPrimaryIP, sUserEmail);
                    }
                    else
                    {
                        bPartResult = false;
                        EventLog.AddLog("Cannot delete project in 60s, test fail");
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog(@"Error occurred when create project and scada: " + ex.ToString());
                    bPartResult = false;
                }
                sw.Stop();
                PrintStep("Create project and scada", "Create project and scada", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());
            }
            Thread.Sleep(1000);
            
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

        private void SetLogDataMaintenance(string sTestLogFolder)
        {
            driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/syslog/ArchPg.asp')]")).Click();

            //DataLog Trend
            driver.FindElement(By.Name("TrendArchChkAll")).Click();
            Thread.Sleep(500);
            driver.FindElement(By.Name("ArchFolder1")).Clear();
            driver.FindElement(By.Name("ArchFolder1")).SendKeys(sTestLogFolder);

            //ODBC Log
            driver.FindElement(By.Name("ODBCArchChkAll")).Click();
            Thread.Sleep(500);
            driver.FindElement(By.Name("ODBCDeleChkAll")).Click();
            Thread.Sleep(500);
            driver.FindElement(By.Name("ODBCExpTimeType1")).Click();
            driver.FindElement(By.Name("ODBCExpTimeType2")).Click();
            driver.FindElement(By.Name("ODBCExpTimeType3")).Click();
            driver.FindElement(By.Name("ODBCExpTimeType4")).Click();
            driver.FindElement(By.Name("ODBCExpTimeType5")).Click();
            driver.FindElement(By.Name("ODBCExpTimeType6")).Click();
            driver.FindElement(By.Name("ODBCExpTimeType7")).Click();
            driver.FindElement(By.Name("ODBCExpTimeType8")).Click();
            driver.FindElement(By.Name("ArchFolder2")).Clear();
            driver.FindElement(By.Name("ArchFolder2")).SendKeys(sTestLogFolder);

            //Excel Report Maintenance
            driver.FindElement(By.Name("ExcelArchChkAll")).Click();
            Thread.Sleep(500);
            driver.FindElement(By.Name("ArchFolder3")).Clear();
            driver.FindElement(By.Name("ArchFolder3")).SendKeys(sTestLogFolder);

            driver.FindElement(By.Name("submit")).Click();
        }

        private void CreateProject(string sProjectName, string sWebAccessIP, string sLanguage)
        {
            EventLog.AddLog("Create a new project");
            // Create a new project
            driver.FindElement(By.Name("ProjName")).Clear();
            driver.FindElement(By.Name("ProjName")).SendKeys(sProjectName);
            driver.FindElement(By.Name("ProjIPLong")).Clear();
            driver.FindElement(By.Name("ProjIPLong")).SendKeys(sWebAccessIP);
            driver.FindElement(By.Name("LogToSystemLog")).Click();
            driver.FindElement(By.Name("submit")).Click();
            Thread.Sleep(1000);

            // Confirm to create
            string alertText = driver.SwitchTo().Alert().Text;
            //string alertText = api.GetAlartTxt();
            switch (sLanguage)
            {
                case "ENG":
                    if (alertText == "Do you want to create a new project ( Project Name : " + sProjectName + " )? ")
                        driver.SwitchTo().Alert().Accept();
                    break;
                case "CHT":
                    if (alertText == "您要建立新的工程 ( 工程名稱 : " + sProjectName + " )? ")
                        driver.SwitchTo().Alert().Accept();
                    break;
                case "CHS":
                    if (alertText == "你想建立新的工程 ( 工程名称 : " + sProjectName + " )? ")
                        driver.SwitchTo().Alert().Accept();
                    break;
                case "JPN":
                    if (alertText == "新しいﾌﾟﾛｼﾞｪｸﾄを作成しますか ( ﾌﾟﾛｼﾞｪｸﾄ名 : " + sProjectName + " )? ")
                        driver.SwitchTo().Alert().Accept();
                    break;
                case "KRN":
                    if (alertText == "새 프로젝트를 생성할까요? ( 프로젝트명 : " + sProjectName + " )? ")
                        driver.SwitchTo().Alert().Accept();
                    break;
                case "FRN":
                    if (alertText == "Voulez-vous créer un nouveau projet ( Nom projet : " + sProjectName + " )? ")
                        driver.SwitchTo().Alert().Accept();
                    break;

                default:
                    if (alertText == "Do you want to create a new project ( Project Name : " + sProjectName + " )? ")
                        driver.SwitchTo().Alert().Accept();
                    break;
            }
        }

        private void CreateSCADANode(string sWebAccessIP, string sUserEmail)
        {
            // Create SCADA node
            driver.SwitchTo().Frame("rightFrame");
            driver.FindElement(By.XPath("//a[2]/font/b")).Click();
            Thread.Sleep(1000);

            EventLog.AddLog("Enter some required info");
            driver.FindElement(By.Name("NodeName")).Clear();
            driver.FindElement(By.Name("NodeName")).SendKeys("TestSCADA");
            driver.FindElement(By.Name("AddressLong")).Clear();
            driver.FindElement(By.Name("AddressLong")).SendKeys(sWebAccessIP);

            driver.FindElement(By.Name("EMAIL_SERVER")).Clear();                         //Outgoing Email (SMTP) Server
            driver.FindElement(By.Name("EMAIL_SERVER")).SendKeys("smtp.mail.yahoo.com");
            driver.FindElement(By.Name("EMAIL_PORT")).Clear();
            driver.FindElement(By.Name("EMAIL_PORT")).SendKeys("587");
            driver.FindElement(By.Name("EMAIL_ADDRESS")).Clear();
            driver.FindElement(By.Name("EMAIL_ADDRESS")).SendKeys("webaccess2016@yahoo.com");
            driver.FindElement(By.Name("EMAIL_USER")).Clear();                         //Email Account Name
            driver.FindElement(By.Name("EMAIL_USER")).SendKeys("webaccess2016");

            driver.FindElement(By.Name("D_EMAIL_PASSWORD")).Clear();
            driver.FindElement(By.Name("D_EMAIL_PASSWORD")).SendKeys("scada2016");
            driver.FindElement(By.Name("D_EMAIL_PASSWORDB")).Clear();
            driver.FindElement(By.Name("D_EMAIL_PASSWORDB")).SendKeys("scada2016");

            driver.FindElement(By.Name("EMAIL_FROM")).Clear();
            driver.FindElement(By.Name("EMAIL_FROM")).SendKeys("webaccess2016@yahoo.com");
            driver.FindElement(By.Name("EMAIL_TO_SRPT")).Clear();
            driver.FindElement(By.Name("EMAIL_TO_SRPT")).SendKeys("webaccess2016@yahoo.com");
            driver.FindElement(By.Name("EMAIL_CC_SRPT")).Clear();
            driver.FindElement(By.Name("EMAIL_CC_SRPT")).SendKeys(sUserEmail);
            driver.FindElement(By.Name("EMAIL_TO")).Clear();
            driver.FindElement(By.Name("EMAIL_TO")).SendKeys("webaccess2016@yahoo.com");
            driver.FindElement(By.Name("EMAIL_CC")).Clear();
            driver.FindElement(By.Name("EMAIL_CC")).SendKeys(sUserEmail);

            driver.FindElement(By.Name("ALARM_LOG_TO_ODBC")).Click();
            driver.FindElement(By.Name("CHANGE_LOG_TO_ODBC")).Click();
            driver.FindElement(By.Name("DATA_LOG_TO_ODBC")).Click();
            driver.FindElement(By.Name("DATA_LOG_USE_RTDB")).Click();
            driver.FindElement(By.Name("Submit")).Click();
            Thread.Sleep(1000);

            EventLog.AddLog("Check if SCADA node exist");
            driver.SwitchTo().ParentFrame();
            driver.SwitchTo().Frame("leftFrame");
            if (!wcf.IsTestElementPresent(driver, "XPath", "//a[contains(@href, '/broadWeb/bwMainRight.asp') and contains(@href, 'name=TestSCADA')]"))
                bPartResult = false;
        }

        private void DeleteProject(string sProjectName, string sLanguage)
        {
            EventLog.AddLog("Delete " + sProjectName + " project");
            driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/project/deleteProject.asp?') and contains(@href, 'ProjName=" + sProjectName + "')]")).Click();
            // Confirm to delete Project
            Thread.Sleep(1000);
            string alertText = driver.SwitchTo().Alert().Text;
            Thread.Sleep(1000);
            switch (sLanguage)
            {
                case "ENG":
                    if (alertText == "Delete this project (" + sProjectName + "), are you sure?")
                        driver.SwitchTo().Alert().Accept();
                    break;
                case "CHT":
                    if (alertText == "您確定要刪除這個工程(" + sProjectName + ")?")
                        driver.SwitchTo().Alert().Accept();
                    break;
                case "CHS":
                    if (alertText == "您肯定要删除工程(" + sProjectName + ")吗?")
                        driver.SwitchTo().Alert().Accept();
                    break;
                case "JPN":
                    if (alertText == "このﾌﾟﾛｼﾞｪｸﾄ (" + sProjectName + ")を削除してもよろしいですか? ")
                        driver.SwitchTo().Alert().Accept();
                    break;
                case "KRN":
                    if (alertText == "이 프로젝트(" + sProjectName + ")를 삭제합니다. 계속하시겠습니까?")
                        driver.SwitchTo().Alert().Accept();
                    break;
                case "FRN":
                    if (alertText == "Supprimer ce projet (" + sProjectName + "), êtes-vous sûr ?")
                        driver.SwitchTo().Alert().Accept();
                    break;

                default:
                    if (alertText == "Delete this project (" + sProjectName + "), are you sure?")
                        driver.SwitchTo().Alert().Accept();
                    break;
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
