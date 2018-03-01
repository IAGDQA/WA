using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using ThirdPartyToolControl;
using iATester;
using CommonFunction;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System.Diagnostics;

namespace CreateExcelReport
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
        string sTestItemName = "CreateExcelReport";
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

            if (bPartResult == true)
            {
                try
                {
                    bPartResult = CreateExcelReport(sPrimaryProject, sUserEmail, sLanguage);
                }
                catch (Exception ex)
                {
                    EventLog.AddLog(@"Error occurred CreateExcelReport: " + ex.ToString());
                    bPartResult = false;
                }
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

        private bool Create_DailyReport_ExcelReport(string sUserEmail, string slanguage)
        {
            try
            {
                for (int t = 1; t <= 4; t++)   /////// template 1~4 ; 8min
                {
                    driver.FindElement(By.XPath("//a[contains(@href, 'addRptCfg.aspx')]")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("rptName")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("rptName")).SendKeys(string.Format("ER_ODBC_T{0}_DR_8min_LastValue_ER", t));
                    Thread.Sleep(1000);

                    IList<IWebElement> oRadioButton = driver.FindElements(By.Name("dataSrc"));
                    int Size = oRadioButton.Count;
                    for (int i = 0; i < Size; i++)
                    {
                        // Store the checkbox name to the string variable, using 'Value' attribute
                        String Value = oRadioButton.ElementAt(i).GetAttribute("value");

                        // Select the checkbox it the value of the checkbox is same what you are looking for
                        if (Value.Equals("1"))
                        {
                            oRadioButton.ElementAt(i).Click();
                            // This will take the execution out of for loop
                            break;
                        }
                    }
                    
                    Thread.Sleep(1000);
                    new SelectElement(driver.FindElement(By.Name("selectTemplate"))).SelectByText(string.Format("template{0}.xlsx", t));
                    Thread.Sleep(1000);
                    
                    switch (slanguage)
                    {
                        case "ENG":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Daily Report");
                            break;
                        case "CHT":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("日報表");
                            break;
                        case "CHS":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("日报表");
                            break;
                        case "JPN":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("日報");
                            break;
                        case "KRN":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("일간 보고");
                            break;
                        case "FRN":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Rapport quotidien");
                            break;

                        default:
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Daily Report");
                            break;
                    }
                    Thread.Sleep(1000);

                    // set end date
                    driver.FindElement(By.Name("tEnd")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("tEnd")).SendKeys("23:59");
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                    Thread.Sleep(1000);

                    driver.FindElement(By.Name("interval")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("interval")).SendKeys("8");                // Time Interval = 8
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//input[@name='fileType'])[2]")).Click();  // Time Unit: Minute
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("//input[@name='valueType']")).Click();      // Data Type: Last Value
                    Thread.Sleep(1000);

                    // 湊到最大上限32個TAG
                    string[] ReportTagName = { "Calc_ConAna", "Calc_ConDis", "Calc_ModBusAI", "Calc_ModBusAO", "Calc_ModBusDI", "Calc_ModBusDO", "Calc_OPCDA", "Calc_OPCUA","Calc_System",
                                         "Calc_ANDConDisAll", "Calc_ANDDIAll", "Calc_ANDDOAll", "Calc_AvgAIAll", "Calc_AvgAOAll", "Calc_AvgConAnaAll", "Calc_AvgSystemAll",
                                         "AT_AI0050", "AT_AI0100", "AT_AI0150", "AT_AO0050", "AT_AO0100", "AT_AO0150", "AT_DI0050", "AT_DI0100", "AT_DI0150", "AT_DO0050",
                                         "AT_DO0100", "AT_DO0150", "ConAna_0100", "ConDis_0100", "SystemSec_0100", "SystemSec_0200"};

                    for (int i = 0; i < ReportTagName.Length; i++)
                    {
                        new SelectElement(driver.FindElement(By.Id("tagsLeftList"))).SelectByText(ReportTagName[i]);
                    }
                    Thread.Sleep(1000);

                    driver.FindElement(By.Id("torightBtn")).Click();
                    Thread.Sleep(1000);
                    new SelectElement(driver.FindElement(By.Id("tagsLeftList"))).SelectByText(ReportTagName[0]);
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//input[@name='attachFormat'])[2]")).Click();   // Send Email: Excel Report     
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("emailto")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("emailto")).SendKeys(sUserEmail);
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("cfgSubmit")).Click();
                }

                for (int t = 1; t <= 4; t++)   ////// template 1~4 ; 1hour
                {
                    driver.FindElement(By.XPath("//a[contains(@href, 'addRptCfg.aspx')]")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("rptName")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("rptName")).SendKeys(string.Format("ER_ODBC_T{0}_DR_1hour_LastValue_ER", t));
                    Thread.Sleep(1000);

                    IList<IWebElement> oRadioButton = driver.FindElements(By.Name("dataSrc"));
                    int Size = oRadioButton.Count;
                    for (int i = 0; i < Size; i++)
                    {
                        // Store the checkbox name to the string variable, using 'Value' attribute
                        String Value = oRadioButton.ElementAt(i).GetAttribute("value");

                        // Select the checkbox it the value of the checkbox is same what you are looking for
                        if (Value.Equals("1"))
                        {
                            oRadioButton.ElementAt(i).Click();
                            // This will take the execution out of for loop
                            break;
                        }
                    }
                    Thread.Sleep(1000);
                    new SelectElement(driver.FindElement(By.Name("selectTemplate"))).SelectByText(string.Format("template{0}.xlsx", t));
                    Thread.Sleep(1000);

                    switch (slanguage)
                    {
                        case "ENG":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Daily Report");
                            break;
                        case "CHT":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("日報表");
                            break;
                        case "CHS":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("日报表");
                            break;
                        case "JPN":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("日報");
                            break;
                        case "KRN":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("일간 보고");
                            break;
                        case "FRN":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Rapport quotidien");
                            break;

                        default:
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Daily Report");
                            break;
                    }
                    Thread.Sleep(1000);

                    // set end date
                    driver.FindElement(By.Name("tEnd")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("tEnd")).SendKeys("23:59");
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                    Thread.Sleep(1000);

                    driver.FindElement(By.Name("interval")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("interval")).SendKeys("1");                // Time Interval = 1
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//input[@name='fileType'])[3]")).Click();  // Time Unit: hour
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("//input[@name='valueType']")).Click();      // Data Type: Last Value
                    Thread.Sleep(1000);

                    string[] ReportTagName = { "Calc_ConAna", "Calc_ConDis", "Calc_ModBusAI", "Calc_ModBusAO", "Calc_ModBusDI", "Calc_ModBusDO", "Calc_OPCDA", "Calc_OPCUA","Calc_System",
                                         "Calc_ANDConDisAll", "Calc_ANDDIAll", "Calc_ANDDOAll", "Calc_AvgAIAll", "Calc_AvgAOAll", "Calc_AvgConAnaAll", "Calc_AvgSystemAll",
                                         "AT_AI0050", "AT_AI0100", "AT_AI0150", "AT_AO0050", "AT_AO0100", "AT_AO0150", "AT_DI0050", "AT_DI0100", "AT_DI0150", "AT_DO0050",
                                         "AT_DO0100", "AT_DO0150", "ConAna_0100", "ConDis_0100", "SystemSec_0100", "SystemSec_0200"};

                    for (int i = 0; i < ReportTagName.Length; i++)
                    {
                        new SelectElement(driver.FindElement(By.Id("tagsLeftList"))).SelectByText(ReportTagName[i]);
                    }
                    Thread.Sleep(1000);

                    driver.FindElement(By.Id("torightBtn")).Click();
                    Thread.Sleep(1000);
                    new SelectElement(driver.FindElement(By.Id("tagsLeftList"))).SelectByText(ReportTagName[0]);
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//input[@name='attachFormat'])[2]")).Click();   // Send Email: Excel Report     
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("emailto")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("emailto")).SendKeys(sUserEmail);
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("cfgSubmit")).Click();
                }

                for (int d = 1; d <= 3; d++)   ////// data type = Maximum or  Minimum or  Average
                {
                    driver.FindElement(By.XPath("//a[contains(@href, 'addRptCfg.aspx')]")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("rptName")).Clear();
                    Thread.Sleep(1000);
                    
                    if (d == 1)
                    {
                        driver.FindElement(By.Name("rptName")).SendKeys("ER_ODBC_T1_DR_1hour_Max_ER");
                    }
                    else if (d == 2)
                    {
                        driver.FindElement(By.Name("rptName")).SendKeys("ER_ODBC_T1_DR_1hour_Min_ER");
                    }
                    else if (d == 3)
                    {
                        driver.FindElement(By.Name("rptName")).SendKeys("ER_ODBC_T1_DR_1hour_Avg_ER");
                    }
                    
                    Thread.Sleep(1000);

                    IList<IWebElement> oRadioButton = driver.FindElements(By.Name("dataSrc"));
                    int Size = oRadioButton.Count;
                    for (int i = 0; i < Size; i++)
                    {
                        // Store the checkbox name to the string variable, using 'Value' attribute
                        String Value = oRadioButton.ElementAt(i).GetAttribute("value");

                        // Select the checkbox it the value of the checkbox is same what you are looking for
                        if (Value.Equals("1"))
                        {
                            oRadioButton.ElementAt(i).Click();
                            // This will take the execution out of for loop
                            break;
                        }
                    }
                    Thread.Sleep(1000);

                    switch (slanguage)
                    {
                        case "ENG":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Daily Report");
                            break;
                        case "CHT":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("日報表");
                            break;
                        case "CHS":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("日报表");
                            break;
                        case "JPN":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("日報");
                            break;
                        case "KRN":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("일간 보고");
                            break;
                        case "FRN":
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Rapport quotidien");
                            break;

                        default:
                            new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Daily Report");
                            break;
                    }
                    Thread.Sleep(1000);

                    // set end date
                    driver.FindElement(By.Name("tEnd")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("tEnd")).SendKeys("23:59");
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                    Thread.Sleep(1000);

                    driver.FindElement(By.Name("interval")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("interval")).SendKeys("1");                // Time Interval = 1
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//input[@name='fileType'])[3]")).Click();  // Time Unit: hour
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath(string.Format("(//input[@name='valueType'])[{0}]", d + 1))).Click();      // Data Type: Last Value
                    Thread.Sleep(1000);

                    // 湊到最大上限32個TAG
                    string[] ReportTagName = { "Calc_ConAna", "Calc_ConDis", "Calc_ModBusAI", "Calc_ModBusAO", "Calc_ModBusDI", "Calc_ModBusDO", "Calc_OPCDA", "Calc_OPCUA","Calc_System",
                                         "Calc_ANDConDisAll", "Calc_ANDDIAll", "Calc_ANDDOAll", "Calc_AvgAIAll", "Calc_AvgAOAll", "Calc_AvgConAnaAll", "Calc_AvgSystemAll",
                                         "AT_AI0050", "AT_AI0100", "AT_AI0150", "AT_AO0050", "AT_AO0100", "AT_AO0150", "AT_DI0050", "AT_DI0100", "AT_DI0150", "AT_DO0050",
                                         "AT_DO0100", "AT_DO0150", "ConAna_0100", "ConDis_0100", "SystemSec_0100", "SystemSec_0200"};

                    for (int i = 0; i < ReportTagName.Length; i++)
                    {
                        new SelectElement(driver.FindElement(By.Id("tagsLeftList"))).SelectByText(ReportTagName[i]);
                    }

                    driver.FindElement(By.Id("torightBtn")).Click();
                    Thread.Sleep(1000);
                    new SelectElement(driver.FindElement(By.Id("tagsLeftList"))).SelectByText(ReportTagName[0]);
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//input[@name='attachFormat'])[2]")).Click();   // Send Email: Excel Report     
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("emailto")).Clear();
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("emailto")).SendKeys(sUserEmail);
                    Thread.Sleep(1000);
                    driver.FindElement(By.Name("cfgSubmit")).Click();
                }
            }
            catch(Exception ex)
            {
                EventLog.AddLog(@"Error occurred Create_DailyReport_ExcelReport: " + ex.ToString());
                return false;
            }
            return true;
        }

        private bool Create_SelfDefined_ExcelReport(string sUserEmail, string sLanguage)
        {
            try
            {
                driver.SwitchTo().Frame("rightFrame");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/WaExlViewer/WaExlViewer.asp')]")).Click();
                Thread.Sleep(1000);

                EventLog.AddLog("Create SelfDefined Data log ExcelReport");
                if (Create_SelfDefined_Datalog_ExcelReport(sUserEmail, sLanguage) == false)
                {
                    return false;
                }

                EventLog.AddLog("Create SelfDefined ODBC Data ExcelReport");
                if (Create_SelfDefined_ODBCData_ExcelReport(sUserEmail, sLanguage) == false)
                {
                    return false;
                }

                EventLog.AddLog("Create SelfDefined Alarm ExcelReport");
                if (Create_SelfDefined_Alarm_ExcelReport(sUserEmail, sLanguage) == false)
                {
                    return false;
                }

                EventLog.AddLog("Create SelfDefined Action Log ExcelReport");
                if (Create_SelfDefined_ActionLog_ExcelReport(sUserEmail, sLanguage) == false)
                {
                    return false;
                }

                EventLog.AddLog("Create SelfDefined Event Log ExcelReport");
                if (Create_SelfDefined_EventLog_ExcelReport(sUserEmail, sLanguage) == false)
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred Create_SelfDefined_ExcelReport: " + ex.ToString());
                return false;
            }
            return true;
        }

        private bool Create_SelfDefined_Datalog_ExcelReport(string sUserEmail, string slanguage)
        {
            try
            {
                driver.FindElement(By.XPath("//a[contains(@href, 'addRptCfg.aspx')]")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("rptName")).Clear();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("rptName")).SendKeys("ExcelReport_SelfDefined_DataLog_Excel");
                Thread.Sleep(1000);
                driver.FindElement(By.Name("dataSrc")).Click();  // Click "Data Log" button
                new SelectElement(driver.FindElement(By.Name("selectTemplate"))).SelectByText("template1.xlsx");
                Thread.Sleep(1000);

                switch (slanguage)
                {
                    case "ENG":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                    case "CHT":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("自訂");
                        break;
                    case "CHS":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("自订");
                        break;
                    case "JPN":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                    case "KRN":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                    case "FRN":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;

                    default:
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                }

                // set today as start/end date
                string sToday = DateTime.Now.ToString("%d");
                string sTomorrow = DateTime.Now.AddDays(+1).ToString("%d");
                driver.FindElement(By.Name("tStart")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.LinkText(sToday)).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                
                if (sTomorrow != "1")
                {
                    driver.FindElement(By.Name("tEnd")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.LinkText(sTomorrow)).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                }
                else  // 跳頁
                {
                    driver.FindElement(By.Name("tEnd")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.CssSelector("span.ui-icon.ui-icon-circle-triangle-e")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.LinkText(sTomorrow)).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                }

                driver.FindElement(By.Name("interval")).Clear();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("interval")).SendKeys("1");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@name='fileType'])[2]")).Click();
                Thread.Sleep(1000);
                
                // 湊到最大上限32個TAG
                string[] ReportTagName = { "Calc_ConAna", "Calc_ConDis", "Calc_ModBusAI", "Calc_ModBusAO", "Calc_ModBusDI", "Calc_ModBusDO", "Calc_OPCDA", "Calc_OPCUA","Calc_System",
                                     "Calc_ANDConDisAll", "Calc_ANDDIAll", "Calc_ANDDOAll", "Calc_AvgAIAll", "Calc_AvgAOAll", "Calc_AvgConAnaAll", "Calc_AvgSystemAll",
                                     "AT_AI0050", "AT_AI0100", "AT_AI0150", "AT_AO0050", "AT_AO0100", "AT_AO0150", "AT_DI0050", "AT_DI0100", "AT_DI0150", "AT_DO0050",
                                     "AT_DO0100", "AT_DO0150", "ConAna_0100", "ConDis_0100", "SystemSec_0100", "SystemSec_0200"};
                for (int i = 0; i < ReportTagName.Length; i++)
                {
                    new SelectElement(driver.FindElement(By.Id("tagsLeftList"))).SelectByText(ReportTagName[i]);
                }

                driver.FindElement(By.Id("torightBtn")).Click();
                Thread.Sleep(1000);
                new SelectElement(driver.FindElement(By.Id("tagsLeftList"))).SelectByText(ReportTagName[0]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@name='attachFormat'])[2]")).Click();   // Send Email: Excel Report     
                Thread.Sleep(1000);
                driver.FindElement(By.Name("emailto")).Clear();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("emailto")).SendKeys(sUserEmail);
                Thread.Sleep(1000);
                driver.FindElement(By.Name("cfgSubmit")).Click();
                
            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred Create_SelfDefined_Datalog_ExcelReport: " + ex.ToString());
                return false;
            }
            return true;
            
        }

        private bool Create_SelfDefined_ODBCData_ExcelReport(string sUserEmail, string slanguage)
        {
            try
            {
                driver.FindElement(By.XPath("//a[contains(@href, 'addRptCfg.aspx')]")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("rptName")).Clear();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("rptName")).SendKeys("ExcelReport_SelfDefined_ODBCData_Excel");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@name='dataSrc'])[2]")).Click();  // ODBC
                new SelectElement(driver.FindElement(By.Name("selectTemplate"))).SelectByText("template1.xlsx");
                Thread.Sleep(1000);

                switch (slanguage)
                {
                    case "ENG":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                    case "CHT":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("自訂");
                        break;
                    case "CHS":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("自订");
                        break;
                    case "JPN":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                    case "KRN":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                    case "FRN":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;

                    default:
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                }

                // set today as start/end date
                string sToday = DateTime.Now.ToString("%d");
                string sTomorrow = DateTime.Now.AddDays(+1).ToString("%d");
                driver.FindElement(By.Name("tStart")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.LinkText(sToday)).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                
                if (sTomorrow != "1")
                {
                    driver.FindElement(By.Name("tEnd")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.LinkText(sTomorrow)).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                }
                else  // 跳頁
                {
                    driver.FindElement(By.Name("tEnd")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.CssSelector("span.ui-icon.ui-icon-circle-triangle-e")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.LinkText(sTomorrow)).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                }

                driver.FindElement(By.Name("interval")).Clear();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("interval")).SendKeys("1");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@name='fileType'])[2]")).Click();
                Thread.Sleep(1000);
                
                // 湊到最大上限32個TAG
                string[] ReportTagName = { "Calc_ConAna", "Calc_ConDis", "Calc_ModBusAI", "Calc_ModBusAO", "Calc_ModBusDI", "Calc_ModBusDO", "Calc_OPCDA", "Calc_OPCUA","Calc_System",
                                     "Calc_ANDConDisAll", "Calc_ANDDIAll", "Calc_ANDDOAll", "Calc_AvgAIAll", "Calc_AvgAOAll", "Calc_AvgConAnaAll", "Calc_AvgSystemAll",
                                     "AT_AI0050", "AT_AI0100", "AT_AI0150", "AT_AO0050", "AT_AO0100", "AT_AO0150", "AT_DI0050", "AT_DI0100", "AT_DI0150", "AT_DO0050",
                                     "AT_DO0100", "AT_DO0150", "ConAna_0100", "ConDis_0100", "SystemSec_0100", "SystemSec_0200"};
                for (int i = 0; i < ReportTagName.Length; i++)
                {
                    new SelectElement(driver.FindElement(By.Id("tagsLeftList"))).SelectByText(ReportTagName[i]);
                }

                driver.FindElement(By.Id("torightBtn")).Click();
                Thread.Sleep(1000);
                new SelectElement(driver.FindElement(By.Id("tagsLeftList"))).SelectByText(ReportTagName[0]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@name='attachFormat'])[2]")).Click();   // Send Email: Excel Report     
                Thread.Sleep(1000);
                driver.FindElement(By.Name("emailto")).Clear();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("emailto")).SendKeys(sUserEmail);
                Thread.Sleep(1000);
                driver.FindElement(By.Name("cfgSubmit")).Click();

            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred Create_SelfDefined_ODBCData_ExcelReport: " + ex.ToString());
                return false;
            }
            return true;
            
        }

        private bool Create_SelfDefined_Alarm_ExcelReport(string sUserEmail,string slanguage)
        {
            try
            {
                driver.FindElement(By.XPath("//a[contains(@href, 'addRptCfg.aspx')]")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("rptName")).Clear();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("rptName")).SendKeys("ExcelReport_SelfDefined_Alarm_Excel");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@name='dataSrc'])[3]")).Click();  // Alarm
                new SelectElement(driver.FindElement(By.Name("selectTemplate"))).SelectByText("AlarmTemplate1_ver.xlsx");
                Thread.Sleep(1000);

                switch (slanguage)
                {
                    case "ENG":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                    case "CHT":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("自訂");
                        break;
                    case "CHS":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("自订");
                        break;
                    case "JPN":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                    case "KRN":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                    case "FRN":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;

                    default:
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                }

                // set today as start/end date
                string sToday = DateTime.Now.ToString("%d");
                string sTomorrow = DateTime.Now.AddDays(+1).ToString("%d");
                driver.FindElement(By.Name("tStart")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.LinkText(sToday)).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                
                if (sTomorrow != "1")
                {
                    driver.FindElement(By.Name("tEnd")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.LinkText(sTomorrow)).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                }
                else  // 跳頁
                {
                    driver.FindElement(By.Name("tEnd")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.CssSelector("span.ui-icon.ui-icon-circle-triangle-e")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.LinkText(sTomorrow)).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                }

                // 湊到最大上限32個TAG
                string[] ReportTagName = { "AT_AI0001", "AT_AO0001", "AT_DI0001", "AT_DO0001", "Calc_ConAna", "Calc_ConDis", "ConDis_0001", "SystemSec_0001",
                                       "ConAna_0001", "ConAna_0010", "ConAna_0020", "ConAna_0030", "ConAna_0040", "ConAna_0050", "ConAna_0060", "ConAna_0070",
                                       "ConAna_0080", "ConAna_0090", "ConAna_0100", "ConAna_0110", "ConAna_0120", "ConAna_0130", "ConAna_0140", "ConAna_0150",
                                       "ConAna_0160", "ConAna_0170","ConAna_0180", "ConAna_0190", "ConAna_0200", "ConAna_0210", "ConAna_0220", "ConAna_0230"};
                for (int i = 0; i < ReportTagName.Length; i++)
                {
                    new SelectElement(driver.FindElement(By.Id("tagsLeftList"))).SelectByText(ReportTagName[i]);
                }

                driver.FindElement(By.Id("torightBtn")).Click();
                Thread.Sleep(1000);
                new SelectElement(driver.FindElement(By.Id("tagsLeftList"))).SelectByText(ReportTagName[0]);
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@name='attachFormat'])[2]")).Click();   // Send Email: Excel Report     
                Thread.Sleep(1000);
                driver.FindElement(By.Name("emailto")).Clear();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("emailto")).SendKeys(sUserEmail);
                Thread.Sleep(1000);
                driver.FindElement(By.Name("cfgSubmit")).Click();

            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred Create_SelfDefined_Alarm_ExcelReport: " + ex.ToString());
                return false;
            }
            return true;
        }

        private bool Create_SelfDefined_ActionLog_ExcelReport(string sUserEmail, string slanguage)
        {
            try
            {
                driver.FindElement(By.XPath("//a[contains(@href, 'addRptCfg.aspx')]")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("rptName")).Clear();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("rptName")).SendKeys("ExcelReport_SelfDefined_ActionLog_Excel");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@name='dataSrc'])[4]")).Click();  // Action Log
                new SelectElement(driver.FindElement(By.Name("selectTemplate"))).SelectByText("ActionTemplate1_ver.xlsx");
                Thread.Sleep(1000);

                switch (slanguage)
                {
                    case "ENG":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                    case "CHT":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("自訂");
                        break;
                    case "CHS":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("自订");
                        break;
                    case "JPN":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                    case "KRN":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                    case "FRN":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;

                    default:
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                }

                // set today as start/end date
                string sToday = DateTime.Now.ToString("%d");
                string sTomorrow = DateTime.Now.AddDays(+1).ToString("%d");
                driver.FindElement(By.Name("tStart")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.LinkText(sToday)).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                
                if (sTomorrow != "1")
                {
                    driver.FindElement(By.Name("tEnd")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.LinkText(sTomorrow)).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                }
                else  // 跳頁
                {
                    driver.FindElement(By.Name("tEnd")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.CssSelector("span.ui-icon.ui-icon-circle-triangle-e")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.LinkText(sTomorrow)).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                }

                driver.FindElement(By.XPath("(//input[@name='attachFormat'])[2]")).Click();   // Send Email: Excel Report     
                Thread.Sleep(1000);
                driver.FindElement(By.Name("emailto")).Clear();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("emailto")).SendKeys(sUserEmail);
                Thread.Sleep(1000);
                driver.FindElement(By.Name("cfgSubmit")).Click();

            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred Create_SelfDefined_ActionLog_ExcelReport: " + ex.ToString());
                return false;
            }
            return true;
        }

        private bool Create_SelfDefined_EventLog_ExcelReport(string sUserEmail,string slanguage)
        {
            try
            {
                driver.FindElement(By.XPath("//a[contains(@href, 'addRptCfg.aspx')]")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("rptName")).Clear();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("rptName")).SendKeys("ExcelReport_SelfDefined_EventLog_Excel");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@name='dataSrc'])[5]")).Click();  // Action Log
                new SelectElement(driver.FindElement(By.Name("selectTemplate"))).SelectByText("EventTemplate1_ver.xlsx");
                Thread.Sleep(1000);

                switch (slanguage)
                {
                    case "ENG":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                    case "CHT":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("自訂");
                        break;
                    case "CHS":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("自订");
                        break;
                    case "JPN":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                    case "KRN":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                    case "FRN":
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;

                    default:
                        new SelectElement(driver.FindElement(By.Name("speTimeFmt"))).SelectByText("Self-Defined");
                        break;
                }

                // set today as start/end date
                string sToday = DateTime.Now.ToString("%d");
                string sTomorrow = DateTime.Now.AddDays(+1).ToString("%d");
                driver.FindElement(By.Name("tStart")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.LinkText(sToday)).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                
                if (sTomorrow != "1")
                {
                    driver.FindElement(By.Name("tEnd")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.LinkText(sTomorrow)).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                }
                else  // 跳頁
                {
                    driver.FindElement(By.Name("tEnd")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.CssSelector("span.ui-icon.ui-icon-circle-triangle-e")).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.LinkText(sTomorrow)).Click();
                    Thread.Sleep(1000);
                    driver.FindElement(By.XPath("(//button[@type='button'])[2]")).Click();
                }

                driver.FindElement(By.XPath("(//input[@name='attachFormat'])[2]")).Click();   // Send Email: Excel Report     
                Thread.Sleep(1000);
                driver.FindElement(By.Name("emailto")).Clear();
                Thread.Sleep(1000);
                driver.FindElement(By.Name("emailto")).SendKeys(sUserEmail);
                Thread.Sleep(1000);
                driver.FindElement(By.Name("cfgSubmit")).Click();
                Thread.Sleep(1000);
                
            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred Create_SelfDefined_EventLog_ExcelReport: " + ex.ToString());
                return false;
            }
            return true;
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

        private bool CreateExcelReport(string sProjectName, string sUserEmail, string sLanguage)
        {
            try
            {
                //EventLog.AddLog("Configure project");
                //driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwMain.asp') and contains(@href, 'ProjName=" + sProjectName + "')]")).Click();
                //Thread.Sleep(1000);

                EventLog.AddLog("check if exist any excel report");
                if (ChceckExcelReport() == false)
                {
                    return false;
                }

                EventLog.AddLog("Create DailyReport Excel Report");
                if (Create_DailyReport_ExcelReport(sUserEmail, sLanguage) == false)
                {
                    return false;
                }

                EventLog.AddLog("Return SCADA Page");
                if (ReturnSCADAPage() == false)
                {
                    return false;
                }

                EventLog.AddLog("Create SelfDefined Excel Report");
                if (Create_SelfDefined_ExcelReport(sUserEmail, sLanguage) == false)
                {
                    return false;
                }


            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred CreateExcelReport: " + ex.ToString());
                return false;
            }
            return true;
            
        }

        private bool ChceckExcelReport()
        {
            try
            {
                driver.SwitchTo().Frame("rightFrame");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("/html/body/table/tbody/tr[1]/td/a[30]/font/b")).Click();
                Thread.Sleep(1000);

                IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                Int64 table_tr_length = (Int64)js.ExecuteScript("return document.getElementsByTagName('tr').length")-3;
                if(table_tr_length == 0)
                    EventLog.AddLog(@"Exist No excel report ");
                else
                {
                    EventLog.AddLog(@"Exist excel report and delete it");
                    for(int i=0; i< table_tr_length; i++)
                    {
                        driver.FindElement(By.XPath("//*[@id='cfgTr_0']/td[7]/a")).Click();
                        Thread.Sleep(1000);
                        IAlert alert = driver.SwitchTo().Alert();
                        alert.Accept();
                        Thread.Sleep(1000);
                    }
                }
            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred ChceckExcelReport: " + ex.ToString());
                return false;
            }
            return true;
        }

        private bool ReturnSCADAPage()
        {
            try
            {
                driver.SwitchTo().ParentFrame();
                Thread.Sleep(1000);
                driver.SwitchTo().Frame("leftFrame");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwMainRight.asp') and contains(@href, 'name=TestSCADA')]")).Click();
                Thread.Sleep(1000);
                driver.SwitchTo().ParentFrame();
                Thread.Sleep(1000);
            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred ReturnSCADAPage: " + ex.ToString());
                return false;
            }
            return true;
        }
    }
}