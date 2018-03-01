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

namespace View_and_Save_AlarmLog
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
        string sTestItemName = "View_and_Save_AlarmLog";
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
                    //driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sPrimaryProject + "')]")).Click();
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
            /*
            EventLog.AddLog("Save data to excel");
            SaveDatatoExcel(sProjectName, sTestLogFolder);
            */
            //AlarmLog Data test
            if (bPartResult == true)
            {
                EventLog.AddLog("AlarmLog Data test");
                sw.Reset(); sw.Start();
                try
                {
                    EventLog.AddLog("Check AlarmLog data...");
                    bPartResult = bAlarmLogCheck(sPrimaryProject, sLanguage);
                }
                catch (Exception ex)
                {
                    EventLog.AddLog(@"Error occurred AlarmLog Data test: " + ex.ToString());
                    bPartResult = false;
                }
                sw.Stop();
                PrintStep("Verify", "AlarmLog Data test", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());

                Thread.Sleep(1000);
            }
            Thread.Sleep(500);
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

        private bool bAlarmLogCheck(string sProjectName, string slanguage)
        {
            bool bCheckAlarm = true;
            //string[] ToBeTestTag = { "AT_AI0001", "AT_AO0001", "AT_DI0001", "AT_DO0001", "Calc_ConAna", "Calc_ConDis", "ConAna_0001", "ConDis_0001", "SystemSec_0001" };
            string[] ToBeTestTag = { "Calc_ConAna", "SystemSec_0001", "AT_AO0001", "AT_AI0001", "ConDis_0001", "ConAna_0001", "ConAna_0125", "ConAna_0250" };
            
            for (int i = 0; i < ToBeTestTag.Length; i++)
            {
                EventLog.AddLog("Go to Alarm log setting page");
                driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/syslog/LogPg.asp') and contains(@href, 'pos=alarm')]")).Click();

                // select project name
                EventLog.AddLog("select project name");
                new SelectElement(driver.FindElement(By.Name("ProjNameSel"))).SelectByText(sProjectName);
                Thread.Sleep(3000);

                // set today as start date
                string sToday = DateTime.Now.ToString("%d");
                driver.FindElement(By.Name("DateStart")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.LinkText(sToday)).Click();
                Thread.Sleep(1000);
                EventLog.AddLog("select start date to today: " + sToday);

                // select one tag to get ODBC data
                EventLog.AddLog("select " + ToBeTestTag[i] + " to get ODBC data");
                driver.FindElement(By.Id("alltags")).Click();

                new SelectElement(driver.FindElement(By.Id("TagNameSel"))).SelectByText(ToBeTestTag[i]);

                driver.FindElement(By.Id("addtag")).Click();
                new SelectElement(driver.FindElement(By.Id("TagNameSelResult"))).SelectByText(ToBeTestTag[i]);

                Thread.Sleep(1000);
                driver.FindElement(By.Name("submit")).Click();
                //PrintStep("Set and get " + ToBeTestTag[i] + " alarm ODBC tag data");
                EventLog.AddLog("Get " + ToBeTestTag[i] + " ODBC data");

                Thread.Sleep(5000); // wait to get ODBC data

                driver.FindElement(By.XPath("//*[@id=\"myTable\"]/thead[1]/tr/th[3]/a")).Click();    // click time to sort data
                Thread.Sleep(5000);
                //api.ByXpath("//*[@id=\"myTable\"]/thead[1]/tr/th[4]/a").Click();    // click tagname to sort data
                //Thread.Sleep(5000);

                bool bRes_ConAna = true;
                if (ToBeTestTag[i] == "ConAna_0001" || ToBeTestTag[i] == "ConAna_0125" || ToBeTestTag[i] == "ConAna_0250")
                    bRes_ConAna = bCheckConAnaRecordAlarm(ToBeTestTag[i], slanguage);

                bool bRes_ConDis = true;
                if (ToBeTestTag[i] == "ConDis_0001")
                    bRes_ConDis = bCheckConDisRecordAlarm(ToBeTestTag[i], slanguage);

                bool bRes_AI = true;
                if (ToBeTestTag[i] == "AT_AI0001")
                    bRes_AI = bCheckAIRecordAlarm(ToBeTestTag[i], slanguage);

                bool bRes_AO = true;
                if (ToBeTestTag[i] == "AT_AO0001")
                    bRes_AO = bCheckAORecordAlarm(ToBeTestTag[i], slanguage);

                bool bRes_Sys = true;
                if (ToBeTestTag[i] == "SystemSec_0001")
                    bRes_Sys = bCheckSysRecordAlarm(ToBeTestTag[i], slanguage);

                bool bRes_Calc = true;
                if (ToBeTestTag[i] == "Calc_ConAna")
                    bRes_Calc = bCheckCalcRecordAlarm(ToBeTestTag[i], slanguage);

                if ((bRes_ConAna && bRes_AI && bRes_AO && bRes_Sys && bRes_Calc) == false)
                {
                    bCheckAlarm = false;
                    break;
                }
                // print screen
                EventLog.PrintScreen(ToBeTestTag[i] + "_AlarmLogData");

                driver.FindElement(By.XPath("//*[@id=\"div1\"]/table/tbody/tr[1]/td[3]/a[5]/font")).Click();     //return to homepage
            }

            return bCheckAlarm;
        }

        private bool bCheckAIRecordAlarm(string sTagName, string slanguage)
        {
            bool bChkTagName = true;
            bool bChkValue = true;

            /////////////////////////////////////////////////////
            // ConAna_0001
            string sRecordTagNameBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[4]")).Text;
            string sRecordTagName = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[4]")).Text;
            string sRecordTagNameAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[4]")).Text;
            EventLog.AddLog(sTagName + " ODBC record TagName(Before): " + sRecordTagNameBefore);
            EventLog.AddLog(sTagName + " ODBC record TagName(Now): " + sRecordTagName);
            EventLog.AddLog(sTagName + " ODBC record TagName(After): " + sRecordTagNameAfter);
            if (sRecordTagNameBefore != sTagName && sRecordTagName != sTagName && sRecordTagNameAfter != sTagName)
            {
                bChkTagName = false;
                EventLog.AddLog(sTagName + " Record TagName check FAIL!!");
            }

            if (bChkTagName)
            {
                string sRecordValueBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[6]")).Text;
                string sRecordValue = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[6]")).Text;
                string sRecordValueAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[6]")).Text;
                EventLog.AddLog(sTagName + " ODBC record value(Before): " + sRecordValueBefore);
                EventLog.AddLog(sTagName + " ODBC record value(Now): " + sRecordValue);
                EventLog.AddLog(sTagName + " ODBC record value(After): " + sRecordValueAfter);

                string sKeyWord = "";
                switch (slanguage)
                {
                    case "ENG":
                        sKeyWord = sTagName+ " - High-High Alarm";
                        break;
                    case "CHT":
                        sKeyWord = sTagName + " - 最高 警報";
                        break;
                    case "CHS":
                        sKeyWord = sTagName + " - 最高 报警";
                        break;
                    case "JPN":
                        sKeyWord = sTagName + " - High-High アラーム";
                        break;
                    case "KRN":
                        sKeyWord = sTagName + " - HH 알람";
                        break;
                    case "FRN":
                        sKeyWord = sTagName + " - Haute-Hau Alarme";
                        break;

                    default:
                        sKeyWord = sTagName + " - High-High Alarm";
                        break;
                }

                if (!sRecordValueBefore.Contains(sKeyWord))
                {
                    EventLog.AddLog("Check " + sTagName + " tag alarm FAIL!!");
                    bChkValue = false;
                }
                else
                {
                    EventLog.AddLog("Check " + sTagName + " tag alarm PASS!!");
                }
            }
            //PrintStep("CheckAIRecordAlarm");
            return bChkTagName && bChkValue;
        }

        private bool bCheckAORecordAlarm(string sTagName, string slanguage)
        {
            bool bChkTagName = true;
            bool bChkValue = true;

            /////////////////////////////////////////////////////
            // ConAna_0001
            string sRecordTagNameBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[4]")).Text;
            string sRecordTagName = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[4]")).Text;
            string sRecordTagNameAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[4]")).Text;
            EventLog.AddLog(sTagName + " ODBC record TagName(Before): " + sRecordTagNameBefore);
            EventLog.AddLog(sTagName + " ODBC record TagName(Now): " + sRecordTagName);
            EventLog.AddLog(sTagName + " ODBC record TagName(After): " + sRecordTagNameAfter);
            if (sRecordTagNameBefore != sTagName && sRecordTagName != sTagName && sRecordTagNameAfter != sTagName)
            {
                bChkTagName = false;
                EventLog.AddLog(sTagName + " Record TagName check FAIL!!");
            }

            if (bChkTagName)
            {
                string sRecordValueBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[6]")).Text;
                string sRecordValue = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[6]")).Text;
                string sRecordValueAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[6]")).Text;
                EventLog.AddLog(sTagName + " ODBC record value(Before): " + sRecordValueBefore);
                EventLog.AddLog(sTagName + " ODBC record value(Now): " + sRecordValue);
                EventLog.AddLog(sTagName + " ODBC record value(After): " + sRecordValueAfter);

                string sKeyWord = "";
                switch (slanguage)
                {
                    case "ENG":
                        sKeyWord = sTagName + " - Low-Low Alarm";
                        break;
                    case "CHT":
                        sKeyWord = sTagName + " - 最低 警報";
                        break;
                    case "CHS":
                        sKeyWord = sTagName + " - 最低 报警";
                        break;
                    case "JPN":
                        sKeyWord = sTagName + " - Low-Low アラーム";
                        break;
                    case "KRN":
                        sKeyWord = sTagName + " - LL 알람";
                        break;
                    case "FRN":
                        sKeyWord = sTagName + " - Basse-Bas Alarme";
                        break;

                    default:
                        sKeyWord = sTagName + " - Low-Low Alarm";
                        break;
                }

                if (!sRecordValueBefore.Contains(sKeyWord))
                {
                    EventLog.AddLog("Check " + sTagName + " tag alarm FAIL!!");
                    bChkValue = false;
                }
                else
                {
                    EventLog.AddLog("Check " + sTagName + " tag alarm PASS!!");
                }
            }
            //PrintStep("CheckAORecordAlarm");
            return bChkTagName && bChkValue;
        }

        private bool bCheckSysRecordAlarm(string sTagName, string slanguage)
        {
            bool bChkTagName = true;
            bool bChkValue = true;

            /////////////////////////////////////////////////////
            // ConAna_0001
            string sRecordTagNameBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[4]")).Text;
            string sRecordTagName = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[4]")).Text;
            string sRecordTagNameAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[4]")).Text;
            EventLog.AddLog(sTagName + " ODBC record TagName(Before): " + sRecordTagNameBefore);
            EventLog.AddLog(sTagName + " ODBC record TagName(Now): " + sRecordTagName);
            EventLog.AddLog(sTagName + " ODBC record TagName(After): " + sRecordTagNameAfter);
            if (sRecordTagNameBefore != sTagName && sRecordTagName != sTagName && sRecordTagNameAfter != sTagName)
            {
                bChkTagName = false;
                EventLog.AddLog(sTagName + " Record TagName check FAIL!!");
            }

            if (bChkTagName)
            {
                string sRecordValueBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[6]")).Text;
                string sRecordValue = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[6]")).Text;
                string sRecordValueAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[6]")).Text;
                EventLog.AddLog(sTagName + " ODBC record value(Before): " + sRecordValueBefore);
                EventLog.AddLog(sTagName + " ODBC record value(Now): " + sRecordValue);
                EventLog.AddLog(sTagName + " ODBC record value(After): " + sRecordValueAfter);

                string sKeyWord = "";
                switch (slanguage)
                {
                    case "ENG":
                        sKeyWord = sTagName + " - High Alarm (59)";
                        break;
                    case "CHT":
                        sKeyWord = sTagName + " - 高的 警報 (59)";
                        break;
                    case "CHS":
                        sKeyWord = sTagName + " - 高 报警 (59)";
                        break;
                    case "JPN":
                        sKeyWord = sTagName + " - High アラーム (59)";
                        break;
                    case "KRN":
                        sKeyWord = sTagName + " - H 알람 (59)";
                        break;
                    case "FRN":
                        sKeyWord = sTagName + " - Haute Alarme (59)";
                        break;

                    default:
                        sKeyWord = sTagName + " - High Alarm (59)";
                        break;
                }

                if (!sRecordValueBefore.Contains(sKeyWord))
                {
                    EventLog.AddLog("Check " + sTagName + " tag alarm FAIL!!");
                    bChkValue = false;
                }
                else
                {
                    EventLog.AddLog("Check " + sTagName + " tag alarm PASS!!");
                }
            }
            //PrintStep("CheckSysRecordAlarm");
            return bChkTagName && bChkValue;
        }

        private bool bCheckCalcRecordAlarm(string sTagName, string slanguage)
        {
            bool bChkTagName = true;
            bool bChkValue = true;

            /////////////////////////////////////////////////////
            // ConAna_0001
            string sRecordTagNameBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[4]")).Text;
            string sRecordTagName = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[4]")).Text;
            string sRecordTagNameAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[4]")).Text;
            EventLog.AddLog(sTagName + " ODBC record TagName(Before): " + sRecordTagNameBefore);
            EventLog.AddLog(sTagName + " ODBC record TagName(Now): " + sRecordTagName);
            EventLog.AddLog(sTagName + " ODBC record TagName(After): " + sRecordTagNameAfter);
            if (sRecordTagNameBefore != sTagName && sRecordTagName != sTagName && sRecordTagNameAfter != sTagName)
            {
                bChkTagName = false;
                EventLog.AddLog(sTagName + " Record TagName check FAIL!!");
            }

            if (bChkTagName)
            {
                string sRecordValueBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[6]")).Text;
                string sRecordValue = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[6]")).Text;
                string sRecordValueAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[6]")).Text;
                EventLog.AddLog(sTagName + " ODBC record value(Before): " + sRecordValueBefore);
                EventLog.AddLog(sTagName + " ODBC record value(Now): " + sRecordValue);
                EventLog.AddLog(sTagName + " ODBC record value(After): " + sRecordValueAfter);

                string sKeyWord = "";
                switch (slanguage)
                {
                    case "ENG":
                        sKeyWord = sTagName + " - Low Alarm";
                        break;
                    case "CHT":
                        sKeyWord = sTagName + " - 低的 警報";
                        break;
                    case "CHS":
                        sKeyWord = sTagName + " - 低 报警";
                        break;
                    case "JPN":
                        sKeyWord = sTagName + " - Low アラーム";
                        break;
                    case "KRN":
                        sKeyWord = sTagName + " - L 알람";
                        break;
                    case "FRN":
                        sKeyWord = sTagName + " - Bas Alarme";
                        break;

                    default:
                        sKeyWord = sTagName + " - Low Alarm";
                        break;
                }

                if (!sRecordValueBefore.Contains(sKeyWord))
                {
                    EventLog.AddLog("Check " + sTagName + " tag alarm FAIL!!");
                    bChkValue = false;
                }
                else
                {
                    EventLog.AddLog("Check " + sTagName + " tag alarm PASS!!");
                }
            }
            //PrintStep("CheckCalcRecordAlarm");
            return bChkTagName && bChkValue;
        }

        private bool bCheckConAnaRecordAlarm(string sTagName, string slanguage)
        {
            bool bChkTagName = true;
            bool bChkValue = true;

            /////////////////////////////////////////////////////
            // ConAna_0001
            string sRecordTagNameBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[4]")).Text;
            string sRecordTagName = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[4]")).Text;
            string sRecordTagNameAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[4]")).Text;
            EventLog.AddLog(sTagName + " ODBC record TagName(Before): " + sRecordTagNameBefore);
            EventLog.AddLog(sTagName + " ODBC record TagName(Now): " + sRecordTagName);
            EventLog.AddLog(sTagName + " ODBC record TagName(After): " + sRecordTagNameAfter);
            if (sRecordTagNameBefore != sTagName && sRecordTagName != sTagName && sRecordTagNameAfter != sTagName)
            {
                bChkTagName = false;
                EventLog.AddLog(sTagName + " Record TagName check FAIL!!");
            }
            
            if (bChkTagName)
            {
                string sRecordValueBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[6]")).Text;
                string sRecordValue = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[6]")).Text;
                string sRecordValueAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[6]")).Text;
                EventLog.AddLog(sTagName + " ODBC record value(Before): " + sRecordValueBefore);
                EventLog.AddLog(sTagName + " ODBC record value(Now): " + sRecordValue);
                EventLog.AddLog(sTagName + " ODBC record value(After): " + sRecordValueAfter);

                string sKeyWord = "";
                switch (slanguage)
                {
                    case "ENG":
                        sKeyWord = sTagName + " - RoC Alarm (51.00)";
                        break;
                    case "CHT":
                        sKeyWord = sTagName + " - 變化率 警報 (51.00)";
                        break;
                    case "CHS":
                        sKeyWord = sTagName + " - 变化率 报警 (51.00)";
                        break;
                    case "JPN":
                        sKeyWord = sTagName + " - RoC アラーム (51.00)";
                        break;
                    case "KRN":
                        sKeyWord = sTagName + " - RoC 알람 (51.00)";
                        break;
                    case "FRN":
                        sKeyWord = sTagName + " - RoC Alarme (51.00)";
                        break;

                    default:
                        sKeyWord = sTagName + " - RoC Alarm (51.00)";
                        break;
                }

                if (sRecordValueBefore != (sKeyWord))
                {
                    EventLog.AddLog("Check "+ sTagName +" tag alarm FAIL!!");
                    bChkValue = false;
                }
                else
                {
                    EventLog.AddLog("Check "+ sTagName +" tag alarm PASS!!");
                }
            }
            //PrintStep("CheckConAnaRecordAlarm");
            return bChkTagName && bChkValue;
        }

        private bool bCheckConDisRecordAlarm(string sTagName, string slanguage)
        {
            bool bChkTagName = true;
            bool bChkValue = true;

            /////////////////////////////////////////////////////
            // ConAna_0001
            string sRecordTagNameBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[4]")).Text;
            string sRecordTagName = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[4]")).Text;
            //string sRecordTagNameAfter = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[3]/td[4]").GetText();
            EventLog.AddLog(sTagName + " ODBC record TagName(Before): " + sRecordTagNameBefore);
            EventLog.AddLog(sTagName + " ODBC record TagName(Now): " + sRecordTagName);
            //EventLog.AddLog(sTagName + " ODBC record TagName(After): " + sRecordTagNameAfter);
            if (sRecordTagNameBefore != sTagName && sRecordTagName != sTagName)
            {
                bChkTagName = false;
                EventLog.AddLog(sTagName + " Record TagName check FAIL!!");
            }

            if (bChkTagName)
            {
                string sRecordValueBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[6]")).Text;
                string sRecordValue = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[6]")).Text;
                //string sRecordValueAfter = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[3]/td[6]").GetText();
                EventLog.AddLog(sTagName + " ODBC record value(Before): " + sRecordValueBefore);
                EventLog.AddLog(sTagName + " ODBC record value(Now): " + sRecordValue);
                //EventLog.AddLog(sTagName + " ODBC record value(After): " + sRecordValueAfter);

                string sKeyWord = "";
                switch (slanguage)
                {
                    case "ENG":
                        sKeyWord = sTagName + " - Discrete Alarm (1)";
                        break;
                    case "CHT":
                        sKeyWord = sTagName + " - 數位量 警報 (1)";
                        break;
                    case "CHS":
                        sKeyWord = sTagName + " - 数字量 报警 (1)";
                        break;
                    case "JPN":
                        sKeyWord = sTagName + " - アラーム (1)";
                        break;
                    case "KRN":
                        sKeyWord = sTagName + " - Discrete 알람 (1)";
                        break;
                    case "FRN":
                        sKeyWord = sTagName + " - Discret Alarme (1)";
                        break;

                    default:
                        sKeyWord = sTagName + " - Discrete Alarm (1)";
                        break;
                }

                if (sRecordValueBefore != (sKeyWord))
                {
                    EventLog.AddLog("Check " + sTagName + " tag alarm FAIL!!");
                    bChkValue = false;
                }
                else
                {
                    EventLog.AddLog("Check " + sTagName + " tag alarm PASS!!");
                }
            }
            //PrintStep("CheckConDisRecordAlarm");
            return bChkTagName && bChkValue;
        }

        private void SaveDatatoExcel(string sProject, string sTestLogFolder)
        {
            string sUserName = Environment.UserName;
            string sourceFile = string.Format(@"C:\Users\{0}\Documents\AlarmLog_Temp.xlsx", sUserName);
            if (System.IO.File.Exists(sourceFile))
                System.IO.File.Delete(sourceFile);

            // Control browser
            int iIE_Handl = tpc.F_FindWindow("IEFrame", "WebAccess Alarm Log - Internet Explorer");
            int iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
            int iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "WebAccess Alarm Log - Internet Explorer");
            int iIE_Handl_4 = tpc.F_FindWindowEx(iIE_Handl_3, 0, "Shell DocObject View", "");
            int iIE_Handl_5 = tpc.F_FindWindowEx(iIE_Handl_4, 0, "Internet Explorer_Server", "");

            if (iIE_Handl_5 > 0)
            {
                int x = 500;
                int y = 500;

                tpc.F_PostMessage(iIE_Handl_5, tpc.V_WM_RBUTTONDOWN, 0, (x & 0xFFFF) + (y & 0xFFFF) * 0x10000);
                //SendMessage(this.Handle, WM_LBUTTONDOWN, 0, (x & 0xFFFF) + (y & 0xFFFF) * 0x10000);
                Thread.Sleep(1000);
                tpc.F_PostMessage(iIE_Handl_5, tpc.V_WM_RBUTTONUP, 0, (x & 0xFFFF) + (y & 0xFFFF) * 0x10000);
                //SendMessage(this.Handle, WM_LBUTTONUP, 0, (x & 0xFFFF) + (y & 0xFFFF) * 0x10000);
                Thread.Sleep(1000);
                // save to excel
                SendKeys.SendWait("X"); // Export to excel
                Thread.Sleep(10000);
            }
            else
            {
                EventLog.AddLog("Cannot get Internet Explorer_Server page handle");
            }

            int iExcel = tpc.F_FindWindow("XLMAIN", "Microsoft Excel - 活頁簿1");
            if (iExcel > 0)                          // 讓開啟的Excel在最上層顯示
            {
                tpc.F_SetForegroundWindow(iExcel);
                Thread.Sleep(5000);
                SendKeys.SendWait("^s");    // save
                Thread.Sleep(2000);
                SendKeys.SendWait("AlarmLog_Temp");
                Thread.Sleep(500);
                SendKeys.SendWait("{ENTER}");
            }
            else
            {
                EventLog.AddLog("Could not find excel handle, excel may not be opened!");
            }

            EventLog.AddLog("Copy AlarmLog_Temp file to Test log folder ");
            string destFile = sTestLogFolder + string.Format("\\AlarmLog_{0:yyyyMMdd_hhmmss}.xlsx", DateTime.Now);
            if (System.IO.File.Exists(sourceFile))
                System.IO.File.Copy(sourceFile, destFile, true);
            else
                EventLog.AddLog(string.Format("The file ( {0} ) is not found.", sourceFile));

            EventLog.AddLog("close excel start");
            Process[] processes = Process.GetProcessesByName("EXCEL");
            foreach (Process p in processes)
            {
                EventLog.AddLog("close excel...");
                p.WaitForExit(2000);
                //p.CloseMainWindow();
                p.Kill();
                p.Close();
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
