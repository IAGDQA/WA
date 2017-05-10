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
using CommonFunction;

namespace View_and_Save_AlarmLog
{
    public partial class Form1 : Form, iATester.iCom
    {
        IAdvSeleniumAPI api;
        cThirdPartyToolControl tpc = new cThirdPartyToolControl();
        cEventLog EventLog = new cEventLog();

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
            long lErrorCode = 0;
            EventLog.AddLog("===View and Save Alarm Log start (by iATester)===");
            if (System.IO.File.Exists(sIniFilePath))    // 再load一次
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===View and Save Alarm Log end (by iATester)===");

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

            api.LinkWebUI(baseUrl + "/broadWeb/bwconfig.asp?username=admin");
            api.ById("userField").Enter("").Submit().Exe();
            PrintStep("Login WebAccess");
            /*
            EventLog.AddLog("Save data to excel");
            SaveDatatoExcel(sProjectName, sTestLogFolder);
            */
            bool bAlmCheck = bAlarmLogCheck(sProjectName);

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

            if (bSeleniumResult && bAlmCheck)
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

        private bool bAlarmLogCheck(string sProjectName)
        {
            bool bCheckAlarm = true;
            //string[] ToBeTestTag = { "AT_AI0001", "AT_AO0001", "AT_DI0001", "AT_DO0001", "Calc_ConAna", "Calc_ConDis", "ConAna_0001", "ConDis_0001", "SystemSec_0001" };
            string[] ToBeTestTag = { "Calc_ConAna", "SystemSec_0001", "AT_AO0001", "AT_AI0001", "ConDis_0001", "ConAna_0001", "ConAna_0125", "ConAna_0250" };
            
            for (int i = 0; i < ToBeTestTag.Length; i++)
            {
                EventLog.AddLog("Go to Alarm log setting page");
                api.ByXpath("//a[contains(@href, '/broadWeb/syslog/LogPg.asp') and contains(@href, 'pos=alarm')]").Click();

                // select project name
                EventLog.AddLog("select project name");
                api.ByName("ProjNameSel").SelectTxt(sProjectName).Exe();
                Thread.Sleep(3000);

                // set today as start date
                string sToday = DateTime.Now.ToString("%d");
                api.ByName("DateStart").Click();
                Thread.Sleep(1000);
                api.ByTxt(sToday).Click();
                Thread.Sleep(1000);
                EventLog.AddLog("select start date to today: " + sToday);

                // select one tag to get ODBC data
                EventLog.AddLog("select " + ToBeTestTag[i] + " to get ODBC data");
                api.ById("alltags").Click();

                api.ById("TagNameSel").SelectTxt(ToBeTestTag[i]).Exe();

                api.ById("addtag").Click();
                api.ById("TagNameSelResult").SelectTxt(ToBeTestTag[i]).Exe();

                Thread.Sleep(1000);
                api.ByName("PageSizeSel").Enter("").Submit().Exe();
                PrintStep("Set and get " + ToBeTestTag[i] + " alarm ODBC tag data");
                EventLog.AddLog("Get " + ToBeTestTag[i] + " ODBC data");

                Thread.Sleep(5000); // wait to get ODBC data

                api.ByXpath("//*[@id=\"myTable\"]/thead[1]/tr/th[3]/a").Click();    // click time to sort data
                Thread.Sleep(5000);
                //api.ByXpath("//*[@id=\"myTable\"]/thead[1]/tr/th[4]/a").Click();    // click tagname to sort data
                //Thread.Sleep(5000);

                bool bRes_ConAna = true;
                if (ToBeTestTag[i] == "ConAna_0001" || ToBeTestTag[i] == "ConAna_0125" || ToBeTestTag[i] == "ConAna_0250")
                    bRes_ConAna = bCheckConAnaRecordAlarm(ToBeTestTag[i]);

                bool bRes_ConDis = true;
                if (ToBeTestTag[i] == "ConDis_0001")
                    bRes_ConDis = bCheckConDisRecordAlarm(ToBeTestTag[i]);

                bool bRes_AI = true;
                if (ToBeTestTag[i] == "AT_AI0001")
                    bRes_AI = bCheckAIRecordAlarm(ToBeTestTag[i]);

                bool bRes_AO = true;
                if (ToBeTestTag[i] == "AT_AO0001")
                    bRes_AO = bCheckAORecordAlarm(ToBeTestTag[i]);

                bool bRes_Sys = true;
                if (ToBeTestTag[i] == "SystemSec_0001")
                    bRes_Sys = bCheckSysRecordAlarm(ToBeTestTag[i]);

                bool bRes_Calc = true;
                if (ToBeTestTag[i] == "Calc_ConAna")
                    bRes_Calc = bCheckCalcRecordAlarm(ToBeTestTag[i]);

                if ((bRes_ConAna && bRes_AI && bRes_AO && bRes_Sys && bRes_Calc) == false)
                {
                    bCheckAlarm = false;
                    break;
                }
                // print screen
                EventLog.PrintScreen(ToBeTestTag[i] + "_AlarmLogData");

                api.ByXpath("//*[@id=\"div1\"]/table/tbody/tr[1]/td[3]/a[5]/font").Click();     //return to homepage
            }

            return bCheckAlarm;
        }

        private bool bCheckAIRecordAlarm(string sTagName)
        {
            bool bChkTagName = true;
            bool bChkValue = true;

            /////////////////////////////////////////////////////
            // ConAna_0001
            string sRecordTagNameBefore = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[1]/td[4]").GetText();
            string sRecordTagName = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[2]/td[4]").GetText();
            string sRecordTagNameAfter = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[3]/td[4]").GetText();
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
                string sRecordValueBefore = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[1]/td[6]").GetText();
                string sRecordValue = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[2]/td[6]").GetText();
                string sRecordValueAfter = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[3]/td[6]").GetText();
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
            PrintStep("CheckAIRecordAlarm");
            return bChkTagName && bChkValue;
        }

        private bool bCheckAORecordAlarm(string sTagName)
        {
            bool bChkTagName = true;
            bool bChkValue = true;

            /////////////////////////////////////////////////////
            // ConAna_0001
            string sRecordTagNameBefore = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[1]/td[4]").GetText();
            string sRecordTagName = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[2]/td[4]").GetText();
            string sRecordTagNameAfter = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[3]/td[4]").GetText();
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
                string sRecordValueBefore = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[1]/td[6]").GetText();
                string sRecordValue = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[2]/td[6]").GetText();
                string sRecordValueAfter = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[3]/td[6]").GetText();
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
            PrintStep("CheckAORecordAlarm");
            return bChkTagName && bChkValue;
        }

        private bool bCheckSysRecordAlarm(string sTagName)
        {
            bool bChkTagName = true;
            bool bChkValue = true;

            /////////////////////////////////////////////////////
            // ConAna_0001
            string sRecordTagNameBefore = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[1]/td[4]").GetText();
            string sRecordTagName = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[2]/td[4]").GetText();
            string sRecordTagNameAfter = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[3]/td[4]").GetText();
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
                string sRecordValueBefore = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[1]/td[6]").GetText();
                string sRecordValue = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[2]/td[6]").GetText();
                string sRecordValueAfter = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[3]/td[6]").GetText();
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
            PrintStep("CheckSysRecordAlarm");
            return bChkTagName && bChkValue;
        }

        private bool bCheckCalcRecordAlarm(string sTagName)
        {
            bool bChkTagName = true;
            bool bChkValue = true;

            /////////////////////////////////////////////////////
            // ConAna_0001
            string sRecordTagNameBefore = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[1]/td[4]").GetText();
            string sRecordTagName = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[2]/td[4]").GetText();
            string sRecordTagNameAfter = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[3]/td[4]").GetText();
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
                string sRecordValueBefore = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[1]/td[6]").GetText();
                string sRecordValue = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[2]/td[6]").GetText();
                string sRecordValueAfter = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[3]/td[6]").GetText();
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
            PrintStep("CheckCalcRecordAlarm");
            return bChkTagName && bChkValue;
        }

        private bool bCheckConAnaRecordAlarm(string sTagName)
        {
            bool bChkTagName = true;
            bool bChkValue = true;

            /////////////////////////////////////////////////////
            // ConAna_0001
            string sRecordTagNameBefore = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[1]/td[4]").GetText();
            string sRecordTagName = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[2]/td[4]").GetText();
            string sRecordTagNameAfter = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[3]/td[4]").GetText();
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
                string sRecordValueBefore = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[1]/td[6]").GetText();
                string sRecordValue = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[2]/td[6]").GetText();
                string sRecordValueAfter = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[3]/td[6]").GetText();
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
            PrintStep("CheckConAnaRecordAlarm");
            return bChkTagName && bChkValue;
        }

        private bool bCheckConDisRecordAlarm(string sTagName)
        {
            bool bChkTagName = true;
            bool bChkValue = true;

            /////////////////////////////////////////////////////
            // ConAna_0001
            string sRecordTagNameBefore = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[1]/td[4]").GetText();
            string sRecordTagName = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[2]/td[4]").GetText();
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
                string sRecordValueBefore = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[1]/td[6]").GetText();
                string sRecordValue = api.ByXpath("//*[@id=\"myTable\"]/tbody/tr[2]/td[6]").GetText();
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
            PrintStep("CheckConDisRecordAlarm");
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
            long lErrorCode = 0;
            EventLog.AddLog("===View and Save Alarm Log start===");
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===View and Save Alarm Log end===");
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
