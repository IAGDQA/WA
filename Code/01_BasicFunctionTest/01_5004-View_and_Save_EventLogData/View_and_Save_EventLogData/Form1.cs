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

namespace View_and_Save_EventLogData
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
        string sTestItemName = "View_and_Save_EventLogData";
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

            //EventLogData test
            if (bPartResult == true)
            {
                EventLog.AddLog("Event LogData test");
                sw.Reset(); sw.Start();
                try
                {
                    if(bPartResult)
                        bPartResult = EventLogTest1(sPrimaryProject);       // 測試只有記錄1個tag的事件
                    Thread.Sleep(3000);
                    if (bPartResult)
                        bPartResult = EventLogTest13579(sPrimaryProject);   // 測試記錄不連續tag的事件(1 3 5 7 9)
                    Thread.Sleep(3000);
                    if (bPartResult)
                        bPartResult = EventLogTestFull(sPrimaryProject);    // 測試記錄80個tag的事件
                }
                catch (Exception ex)
                {
                    EventLog.AddLog(@"Error occurred Event LogData test: " + ex.ToString());
                    bPartResult = false;
                }
                sw.Stop();
                PrintStep("Verify", "Event LogData test", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());

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

        private bool EventLogTestFull(string sProjectName)
        {
            EventLog.AddLog("Go to Event log setting page");
            driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/syslog/LogPg.asp?pos=event')]")).Click();

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

            // select event log name
            new SelectElement(driver.FindElement(By.Name("EveNameSel"))).SelectByText("EventLog_" + sProjectName);
            Thread.Sleep(1000);
            driver.FindElement(By.Name("submit")).Click();
            //PrintStep("Set and get Event Log data");
            EventLog.AddLog("Get Event Log data");

            Thread.Sleep(10000); // wait to get ODBC data

            // print screen
            string fileNameTar = string.Format("EventLogData_{0:yyyyMMdd_hhmmss}", DateTime.Now);
            EventLog.PrintScreen(fileNameTar);

            EventLog.AddLog("Check event log data for 80 continuous log value");
            bool bCheckResult = CheckEventLogData();
            //PrintStep("CheckEventLogData");

            // return to home page
            driver.FindElement(By.XPath("//a[5]/font")).Click();

            return bCheckResult;
        }

        private bool EventLogTest1(string sProjectName)
        {
            EventLog.AddLog("Go to Event log setting page");
            driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/syslog/LogPg.asp?pos=event')]")).Click();

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

            // select event log name
            new SelectElement(driver.FindElement(By.Name("EveNameSel"))).SelectByText("EventLog1_" + sProjectName);
            Thread.Sleep(1000);
            driver.FindElement(By.Name("submit")).Click();
            //PrintStep("Set and get Event Log data");
            EventLog.AddLog("Get Event Log data");

            Thread.Sleep(10000); // wait to get ODBC data

            // print screen
            string fileNameTar = string.Format("EventLogData_{0:yyyyMMdd_hhmmss}", DateTime.Now);
            EventLog.PrintScreen(fileNameTar);
            /*
            EventLog.AddLog("Save data to excel");
            SaveDatatoExcel(sProjectName, sTestLogFolder);
            */
            EventLog.AddLog("Check event log data for 1 log value");
            bool bCheckResult = CheckEventLogData1();
            //PrintStep("CheckEventLogData1");

            // return to home page
            driver.FindElement(By.XPath("//a[5]/font")).Click();

            return bCheckResult;
        }

        private bool EventLogTest13579(string sProjectName)
        {
            EventLog.AddLog("Go to Event log setting page");
            driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/syslog/LogPg.asp?pos=event')]")).Click();

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

            // select event log name
            new SelectElement(driver.FindElement(By.Name("EveNameSel"))).SelectByText("EventLog13579_" + sProjectName);
            Thread.Sleep(1000);
            driver.FindElement(By.Name("submit")).Click();
            //PrintStep("Set and get Event Log data");
            EventLog.AddLog("Get Event Log data");

            Thread.Sleep(10000); // wait to get ODBC data

            // print screen
            string fileNameTar = string.Format("EventLogData_{0:yyyyMMdd_hhmmss}", DateTime.Now);
            EventLog.PrintScreen(fileNameTar);

            EventLog.AddLog("Check event log data for 5 discontinuous log value");
            bool bCheckResult = CheckEventLogData13579();
            //PrintStep("CheckEventLogData13579");

            // return to home page
            driver.FindElement(By.XPath("//a[5]/font")).Click();

            return bCheckResult;
        }

        private bool CheckEventLogData()
        {
            bool bCheckEventLogData = true;
            string sDate = DateTime.Now.ToString("yyyy/M/d");
            //string sDate2 = DateTime.Now.ToString("yyyy-M-d");
            //string sDate3 = DateTime.Now.ToString("yyyy-MM-dd");
            //string sDate4 = DateTime.Now.ToString("dd/MM/yyyy");    // FRN

            string sEventRecordDate = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[1]/font")).Text;

            EventLog.AddLog("Event record date: " + sEventRecordDate);
            EventLog.AddLog("Today is: " + sDate);
            //EventLog.AddLog("Today is: " + sDate2);
            //EventLog.AddLog("Today is: " + sDate3);
            //EventLog.AddLog("Today is: " + sDate4);
            //if ( (sDate == sEventRecordDate) || (sDate2 == sEventRecordDate) || (sDate3 == sEventRecordDate) || (sDate4 == sEventRecordDate))
            //{
            //    EventLog.AddLog("Event record date check PASS!!");
            //}
            //else
            //{
            //    EventLog.AddLog("Event record date check FAIL!!");
            //    bCheckEventLogData = false;
            //}

            driver.FindElement(By.XPath("//*[@id=\"myTable\"]/thead[1]/tr/th[2]/a")).Click();    // click time to sort data
            Thread.Sleep(10000);
            //api.ByXpath("//*[@id=\"myTable\"]/thead[1]/tr/th[3]/a").Click();    // click tagname to sort data
            //Thread.Sleep(10000);

            if (bCheckEventLogData) // 確認記錄事件的名稱是否前1秒後1秒
            {
                string sRecordTimeBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[2]")).Text;
                string sRecordTime = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[2]")).Text;
                string sRecordTimeAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[2]")).Text;
                string sRecordTimeMSBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[3]")).Text;
                string sRecordMSTime = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[3]")).Text;
                string sRecordMSTimeAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[3]")).Text;
                EventLog.AddLog("Event record time(Before): " + sRecordTimeBefore);
                EventLog.AddLog("Event record time(Now): " + sRecordTime);
                EventLog.AddLog("Event record time(After): " + sRecordTimeAfter);

                string[] sBefore_tmp = sRecordTimeBefore.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                string[] sNow_tmp = sRecordTime.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                string[] sAfter_tmp = sRecordTimeAfter.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                if (sRecordTimeBefore != "" && sRecordTime != "" && sRecordTimeAfter != "")
                {
                    if (Int32.Parse(sNow_tmp[2]) - Int32.Parse(sBefore_tmp[2]) == 1 &&      // 確認記錄事件的名稱是否前1秒後1秒
                        Int32.Parse(sAfter_tmp[2]) - Int32.Parse(sNow_tmp[2]) == 1)
                    {
                        EventLog.AddLog("Record time interval check PASS!!");
                    }
                    else if (Int32.Parse(sNow_tmp[2]) - Int32.Parse(sBefore_tmp[2]) == -59 &&      // 59-0-1
                        Int32.Parse(sAfter_tmp[2]) - Int32.Parse(sNow_tmp[2]) == 1)
                    {
                        EventLog.AddLog("Record time interval check PASS!!");
                    }
                    else if (Int32.Parse(sNow_tmp[2]) - Int32.Parse(sBefore_tmp[2]) == 1 &&      // 58-59-0
                        Int32.Parse(sAfter_tmp[2]) - Int32.Parse(sNow_tmp[2]) == -59)
                    {
                        EventLog.AddLog("Record time interval check PASS!!");
                    }
                    else if (Int32.Parse(sNow_tmp[2]) - Int32.Parse(sBefore_tmp[2]) == 0)       // 前後值相等的情況
                    {
                        EventLog.AddLog("Event record ms time(Before): " + sRecordTimeMSBefore);
                        EventLog.AddLog("Event record ms time(Now): " + sRecordMSTime);

                        try
                        {
                            if (Convert.ToDouble(sRecordTimeMSBefore) - Convert.ToDouble(sRecordMSTime) > 500)
                            {
                                EventLog.AddLog("Record time interval check PASS!!");
                            }
                            else
                            {
                                bCheckEventLogData = false;
                                EventLog.AddLog("Record time interval check FAIL!!");
                            }
                        }
                        catch (FormatException)
                        {
                            EventLog.AddLog("Unable to convert record time to a Double.");
                            bPartResult = false;
                        }
                        catch (OverflowException)
                        {
                            EventLog.AddLog(" Record time is outside the range of a Double.");
                            bPartResult = false;
                        }
                    }
                    else if(Int32.Parse(sAfter_tmp[2]) - Int32.Parse(sNow_tmp[2]) == 0) // 前後值相等的情況
                    {
                        EventLog.AddLog("Event record ms time(Now): " + sRecordMSTime);
                        EventLog.AddLog("Event record ms time(After): " + sRecordMSTimeAfter);
                        try
                        {
                            if (Convert.ToDouble(sRecordMSTime) - Convert.ToDouble(sRecordMSTimeAfter) > 900)
                            {
                                EventLog.AddLog("Record time interval check PASS!!");
                            }
                            else
                            {
                                bCheckEventLogData = false;
                                EventLog.AddLog("Record time interval check FAIL!!");
                            }
                        }
                        catch (FormatException)
                        {
                            EventLog.AddLog("Unable to convert record time to a Double.");
                            bPartResult = false;
                        }
                        catch (OverflowException)
                        {
                            EventLog.AddLog(" Record time is outside the range of a Double.");
                            bPartResult = false;
                        }
                    }
                    else
                    {
                        bCheckEventLogData = false;
                        EventLog.AddLog("Record time interval check FAIL!!");
                    }
                }
                else
                {
                    bCheckEventLogData = false;
                    EventLog.AddLog("Record time interval check FAIL!!");
                }
            }

            if (bCheckEventLogData) // 確認記錄的數值是否為51
            {
                for (int i = 1; i <= 237; i=i+3)    //20180202 修正為只測試80個tag
                {
                    string sTagName = driver.FindElement(By.XPath(string.Format("//*[@id=\"myTable\"]/thead[1]/tr/th[{0}]/a", i + 3))).Text;
                    //string sValueBefore = api.ByXpath(string.Format("//*[@id=\"myTable\"]/tbody/tr[1]/td[{0}]", i + 3)).GetText();
                    string sValue = driver.FindElement(By.XPath(string.Format("//*[@id=\"myTable\"]/tbody/tr[2]/td[{0}]/font", i + 3))).Text;
                    //string sValueAfter = api.ByXpath(string.Format("//*[@id=\"myTable\"]/tbody/tr[3]/td[{0}]/font", i + 3)).GetText();

                    //EventLog.AddLog("TagName: " + sTagName + " BeforeValue: " + sValueBefore + " Value: " + sValue + " AfterValue: " + sValueAfter);
                    EventLog.AddLog("TagName: " + sTagName + " Value: " + sValue);
                    /*
                    double number;
                    if (!Double.TryParse(sValue, out number) || number != 51)    // if string to double success!!
                    {
                        EventLog.AddLog("Event log value check FAIL!!");
                        bCheckEventLogData = false;
                        break;
                    }
                    */
                    if (sValue.Trim() != "51.00" && sValue.Trim() != "51")    // if string to double success!!
                    {
                        EventLog.AddLog("Event log value check FAIL!!");
                        bCheckEventLogData = false;
                        break;
                    }
                }
                if (bCheckEventLogData)
                    EventLog.AddLog("Event log value check PASS!!");
            }
            return bCheckEventLogData;
        }  // 測試記錄80個tag的事件

        private bool CheckEventLogData1()
        {
            bool bCheckEventLogData = true;
            string sDate = DateTime.Now.ToString("yyyy/M/d");

            string sEventRecordDate = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[1]/font")).Text;

            EventLog.AddLog("Event record date: " + sEventRecordDate);
            EventLog.AddLog("Today is: " + sDate);

            driver.FindElement(By.XPath("//*[@id=\"myTable\"]/thead[1]/tr/th[2]/a")).Click();    // click time to sort data
            Thread.Sleep(10000);

            if (bCheckEventLogData) // 確認記錄事件的名稱是否前1秒後1秒
            {
                string sRecordTimeBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[2]")).Text;
                string sRecordTime = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[2]")).Text;
                string sRecordTimeAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[2]")).Text;
                string sRecordTimeMSBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[3]")).Text;
                string sRecordMSTime = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[3]")).Text;
                string sRecordMSTimeAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[3]")).Text;
                EventLog.AddLog("Event record time(Before): " + sRecordTimeBefore);
                EventLog.AddLog("Event record time(Now): " + sRecordTime);
                EventLog.AddLog("Event record time(After): " + sRecordTimeAfter);

                string[] sBefore_tmp = sRecordTimeBefore.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                string[] sNow_tmp = sRecordTime.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                string[] sAfter_tmp = sRecordTimeAfter.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                if (sRecordTimeBefore != "" && sRecordTime != "" && sRecordTimeAfter != "")
                {
                    if (Int32.Parse(sNow_tmp[2]) - Int32.Parse(sBefore_tmp[2]) == 1 &&      // 確認記錄事件的名稱是否前1秒後1秒
                        Int32.Parse(sAfter_tmp[2]) - Int32.Parse(sNow_tmp[2]) == 1)
                    {
                        EventLog.AddLog("Record time interval check PASS!!");
                    }
                    else if (Int32.Parse(sNow_tmp[2]) - Int32.Parse(sBefore_tmp[2]) == -59 &&      // 59-0-1
                        Int32.Parse(sAfter_tmp[2]) - Int32.Parse(sNow_tmp[2]) == 1)
                    {
                        EventLog.AddLog("Record time interval check PASS!!");
                    }
                    else if (Int32.Parse(sNow_tmp[2]) - Int32.Parse(sBefore_tmp[2]) == 1 &&      // 58-59-0
                        Int32.Parse(sAfter_tmp[2]) - Int32.Parse(sNow_tmp[2]) == -59)
                    {
                        EventLog.AddLog("Record time interval check PASS!!");
                    }
                    else if (Int32.Parse(sNow_tmp[2]) - Int32.Parse(sBefore_tmp[2]) == 0)       // 前後值相等的情況
                    {
                        EventLog.AddLog("Event record ms time(Before): " + sRecordTimeMSBefore);
                        EventLog.AddLog("Event record ms time(Now): " + sRecordMSTime);
                        try
                        {
                            if (Convert.ToDouble(sRecordTimeMSBefore) - Convert.ToDouble(sRecordMSTime) > 500)
                            {
                                EventLog.AddLog("Record time interval check PASS!!");
                            }
                            else
                            {
                                bCheckEventLogData = false;
                                EventLog.AddLog("Record time interval check FAIL!!");
                            }
                        }
                        catch (FormatException)
                        {
                            EventLog.AddLog("Unable to convert record time to a Double.");
                            bPartResult = false;
                        }
                        catch (OverflowException)
                        {
                            EventLog.AddLog(" Record time is outside the range of a Double.");
                            bPartResult = false;
                        }
                    }
                    else if (Int32.Parse(sAfter_tmp[2]) - Int32.Parse(sNow_tmp[2]) == 0) // 前後值相等的情況
                    {
                        EventLog.AddLog("Event record ms time(Now): " + sRecordMSTime);
                        EventLog.AddLog("Event record ms time(After): " + sRecordMSTimeAfter);

                        try
                        {
                            if (Convert.ToDouble(sRecordMSTime) - Convert.ToDouble(sRecordMSTimeAfter) > 500)
                            {
                                EventLog.AddLog("Record time interval check PASS!!");
                            }
                            else
                            {
                                bCheckEventLogData = false;
                                EventLog.AddLog("Record time interval check FAIL!!");
                            }
                        }
                        catch (FormatException)
                        {
                            EventLog.AddLog("Unable to convert record time to a Double.");
                            bPartResult = false;
                        }
                        catch (OverflowException)
                        {
                            EventLog.AddLog(" Record time is outside the range of a Double.");
                            bPartResult = false;
                        }
                    }
                    else
                    {
                        bCheckEventLogData = false;
                        EventLog.AddLog("Record time interval check FAIL!!");
                    }
                }
                else
                {
                    bCheckEventLogData = false;
                    EventLog.AddLog("Record time interval check FAIL!!");
                }
            }

            if (bCheckEventLogData) // 確認記錄的數值是否為51
            {
                for (int i = 1; i <= 1; i++)
                {
                    string sTagName = driver.FindElement(By.XPath(string.Format("//*[@id=\"myTable\"]/thead[1]/tr/th[{0}]/a", i + 3))).Text;
                    //string sValueBefore = api.ByXpath(string.Format("//*[@id=\"myTable\"]/tbody/tr[1]/td[{0}]", i + 3)).GetText();
                    string sValue = driver.FindElement(By.XPath(string.Format("//*[@id=\"myTable\"]/tbody/tr[2]/td[{0}]/font", i + 3))).Text;
                    //string sValueAfter = api.ByXpath(string.Format("//*[@id=\"myTable\"]/tbody/tr[3]/td[{0}]/font", i + 3)).GetText();

                    //EventLog.AddLog("TagName: " + sTagName + " BeforeValue: " + sValueBefore + " Value: " + sValue + " AfterValue: " + sValueAfter);
                    EventLog.AddLog("TagName: " + sTagName + " Value: " + sValue);
                    /*
                    double number;

                    if (!Double.TryParse(sValue, out number) || number != 51)    // if string to double success!!
                    {                                                            // 此部分不知為何在法文版本一定會進到這個if, 故停用這方式
                        EventLog.AddLog("Event log value check FAIL!!");
                        bCheckEventLogData = false;
                        break;
                    }
                    */
                    if ( sValue.Trim() != "51.00" && sValue.Trim() != "51")    // if string to double success!!
                    {
                        EventLog.AddLog("Event log value check FAIL!!");
                        bCheckEventLogData = false;
                        break;
                    }

                }
                if (bCheckEventLogData)
                    EventLog.AddLog("Event log value check PASS!!");
            }
            return bCheckEventLogData;
        }   // 測試只有記錄1個tag的事件

        private bool CheckEventLogData13579()
        {
            bool bCheckEventLogData = true;
            string sDate = DateTime.Now.ToString("yyyy/M/d");

            string sEventRecordDate = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[1]/font")).Text;

            EventLog.AddLog("Event record date: " + sEventRecordDate);
            EventLog.AddLog("Today is: " + sDate);

            driver.FindElement(By.XPath("//*[@id=\"myTable\"]/thead[1]/tr/th[2]/a")).Click();    // click time to sort data
            Thread.Sleep(10000);

            if (bCheckEventLogData) // 確認記錄事件的名稱是否前1秒後1秒
            {
                string sRecordTimeBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[2]")).Text;
                string sRecordTime = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[2]")).Text;
                string sRecordTimeAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[2]")).Text;
                string sRecordTimeMSBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[3]")).Text;
                string sRecordMSTime = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[3]")).Text;
                string sRecordMSTimeAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[3]")).Text;
                EventLog.AddLog("Event record time(Before): " + sRecordTimeBefore);
                EventLog.AddLog("Event record time(Now): " + sRecordTime);
                EventLog.AddLog("Event record time(After): " + sRecordTimeAfter);

                string[] sBefore_tmp = sRecordTimeBefore.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                string[] sNow_tmp = sRecordTime.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                string[] sAfter_tmp = sRecordTimeAfter.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
                if (sRecordTimeBefore != "" && sRecordTime != "" && sRecordTimeAfter != "")
                {
                    if (Int32.Parse(sNow_tmp[2]) - Int32.Parse(sBefore_tmp[2]) == 1 &&      // 確認記錄事件的名稱是否前1秒後1秒
                        Int32.Parse(sAfter_tmp[2]) - Int32.Parse(sNow_tmp[2]) == 1)
                    {
                        EventLog.AddLog("Record time interval check PASS!!");
                    }
                    else if (Int32.Parse(sNow_tmp[2]) - Int32.Parse(sBefore_tmp[2]) == -59 &&      // 59-0-1
                        Int32.Parse(sAfter_tmp[2]) - Int32.Parse(sNow_tmp[2]) == 1)
                    {
                        EventLog.AddLog("Record time interval check PASS!!");
                    }
                    else if (Int32.Parse(sNow_tmp[2]) - Int32.Parse(sBefore_tmp[2]) == 1 &&      // 58-59-0
                        Int32.Parse(sAfter_tmp[2]) - Int32.Parse(sNow_tmp[2]) == -59)
                    {
                        EventLog.AddLog("Record time interval check PASS!!");
                    }
                    else if (Int32.Parse(sNow_tmp[2]) - Int32.Parse(sBefore_tmp[2]) == 0)       // 前後值相等的情況
                    {
                        EventLog.AddLog("Event record ms time(Before): " + sRecordTimeMSBefore);
                        EventLog.AddLog("Event record ms time(Now): " + sRecordMSTime);
                        try
                        {
                            if (Convert.ToDouble(sRecordTimeMSBefore) - Convert.ToDouble(sRecordMSTime) > 500)
                            {
                                EventLog.AddLog("Record time interval check PASS!!");
                            }
                            else
                            {
                                bCheckEventLogData = false;
                                EventLog.AddLog("Record time interval check FAIL!!");
                            }
                        }
                        catch (FormatException)
                        {
                            EventLog.AddLog("Unable to convert record time to a Double.");
                            bPartResult = false;
                        }
                        catch (OverflowException)
                        {
                            EventLog.AddLog(" Record time is outside the range of a Double.");
                            bPartResult = false;
                        }
                    }
                    else if (Int32.Parse(sAfter_tmp[2]) - Int32.Parse(sNow_tmp[2]) == 0) // 前後值相等的情況
                    {
                        EventLog.AddLog("Event record ms time(Now): " + sRecordMSTime);
                        EventLog.AddLog("Event record ms time(After): " + sRecordMSTimeAfter);
                        try
                        {
                            if (Convert.ToDouble(sRecordMSTime) - Convert.ToDouble(sRecordMSTimeAfter) > 900)
                            {
                                EventLog.AddLog("Record time interval check PASS!!");
                            }
                            else
                            {
                                bCheckEventLogData = false;
                                EventLog.AddLog("Record time interval check FAIL!!");
                            }
                        }
                        catch (FormatException)
                        {
                            EventLog.AddLog("Unable to convert record time to a Double.");
                            bPartResult = false;
                        }
                        catch (OverflowException)
                        {
                            EventLog.AddLog(" Record time is outside the range of a Double.");
                            bPartResult = false;
                        }
                    }
                    else
                    {
                        bCheckEventLogData = false;
                        EventLog.AddLog("Record time interval check FAIL!!");
                    }
                }
                else
                {
                    bCheckEventLogData = false;
                    EventLog.AddLog("Record time interval check FAIL!!");
                }
            }

            if (bCheckEventLogData) // 確認記錄的數值是否為51
            {
                for (int i = 1; i <= 5; i++)
                {
                    string sTagName = driver.FindElement(By.XPath(string.Format("//*[@id=\"myTable\"]/thead[1]/tr/th[{0}]/a", i + 3))).Text;
                    //string sValueBefore = api.ByXpath(string.Format("//*[@id=\"myTable\"]/tbody/tr[1]/td[{0}]", i + 3)).GetText();
                    string sValue = driver.FindElement(By.XPath(string.Format("//*[@id=\"myTable\"]/tbody/tr[2]/td[{0}]/font", i + 3))).Text;
                    //string sValueAfter = api.ByXpath(string.Format("//*[@id=\"myTable\"]/tbody/tr[3]/td[{0}]/font", i + 3)).GetText();

                    //EventLog.AddLog("TagName: " + sTagName + " BeforeValue: " + sValueBefore + " Value: " + sValue + " AfterValue: " + sValueAfter);
                    EventLog.AddLog("TagName: " + sTagName + " Value: " + sValue);
                    /*
                    double number;
                    if (!Double.TryParse(sValue, out number) || number != 51)    // if string to double success!!
                    {
                        EventLog.AddLog("Event log value check FAIL!!");
                        bCheckEventLogData = false;
                        break;
                    }
                    */
                    if (sValue.Trim() != "51.00" && sValue.Trim() != "51")    // if string to double success!!
                    {
                        EventLog.AddLog("Event log value check FAIL!!");
                        bCheckEventLogData = false;
                        break;
                    }
                }
                if (bCheckEventLogData)
                    EventLog.AddLog("Event log value check PASS!!");
            }
            return bCheckEventLogData;
        }  // 測試記錄不連續tag的事件(1 3 5 7 9)

        private void SaveDatatoExcel(string sProject, string sTestLogFolder)
        {
            string sUserName = Environment.UserName;
            string sourceFile = string.Format(@"C:\Users\{0}\Documents\EventLog_Temp.xlsx", sUserName);
            if (System.IO.File.Exists(sourceFile))
                System.IO.File.Delete(sourceFile);

            // Control browser
            int iIE_Handl = tpc.F_FindWindow("IEFrame", "Event Log - Internet Explorer");
            int iIE_Handl_2 = tpc.F_FindWindowEx(iIE_Handl, 0, "Frame Tab", "");
            int iIE_Handl_3 = tpc.F_FindWindowEx(iIE_Handl_2, 0, "TabWindowClass", "Event Log - Internet Explorer");
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
            if(iExcel > 0)                          // 讓開啟的Excel在最上層顯示
            {
                tpc.F_SetForegroundWindow(iExcel);
                Thread.Sleep(5000);
                SendKeys.SendWait("^s");    // save
                Thread.Sleep(2000);
                SendKeys.SendWait("EventLog_Temp");
                Thread.Sleep(500);
                SendKeys.SendWait("{ENTER}");
            }
            else
            {
                EventLog.AddLog("Could not find excel handle, excel may not be opened!");
            }

            EventLog.AddLog("Copy EventLog_Temp file to Test log folder ");
            string destFile = sTestLogFolder + string.Format("\\EventLog_{0:yyyyMMdd_hhmmss}.xlsx", DateTime.Now);
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
