using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows.Forms;
using iATester;
using CommonFunction;
using ThirdPartyToolControl;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using System.Threading;
using OpenQA.Selenium.Support.UI;

namespace AnalogChangeLog_Test
{
    public partial class Form1 : Form, iATester.iCom
    {
        cThirdPartyToolControl tpc = new cThirdPartyToolControl();
        cEventLog EventLog = new cEventLog();
        Stopwatch sw = new Stopwatch();
        private IWebDriver driver;
        bool bFinalResult = true;
        bool bPartResult = true;
        string sTestItemName = "AnalogChangeLog_Test";

        private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
        private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
        internal const int Max_Rows_Val = 65535;
        string baseUrl;
        string sIniFilePath = @"C:\WebAccessAutoTestSetting.ini";

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
            EventLog.AddLog(string.Format("==={0} test start (by iATester)===", sTestItemName));
            if (System.IO.File.Exists(sIniFilePath))    // 再load一次
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog(string.Format("==={0} test end (by iATester)===", sTestItemName));

            if (lErrorCode == 0)
                eResult(this, new ResultEventArgs(iResult.Pass));
            else
                eResult(this, new ResultEventArgs(iResult.Fail));

            eStatus(this, new StatusEventArgs(iStatus.Completion));
        }

        public Form1()
        {
            InitializeComponent();
            Browser.SelectedIndex = 0;
            Text = string.Format("Advantech WebAccess Auto Test ( {0} )", sTestItemName);
            try
            {
                m_DataGridViewCtrlAddDataRow = new DataGridViewCtrlAddDataRow(DataGridViewCtrlAddNewRow);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

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
                InternetExplorerOptions options = new InternetExplorerOptions();
                options.IgnoreZoomLevel = true;
                driver = new InternetExplorerDriver(options);
            }
            else
            {
                EventLog.AddLog("Not support temporary");
                ///driver = new FirefoxDriver();
            }

            /*Login test*/
            sw.Reset(); sw.Start(); bPartResult = true;
            try
            {
                driver.Navigate().GoToUrl(baseUrl + "/broadWeb/bwRoot.asp?username=admin");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwconfig.asp?username=admin')]")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Id("userField")).Submit();
                Thread.Sleep(1000);
            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred logging on: " + ex.ToString());
                bPartResult = false;
            }
            sw.Stop();
            PrintStep("Login", "login Project Manager page", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());
            /*Login test*/

            /*
            Start to write your code here.
            */
            sw.Reset(); sw.Start(); bPartResult = true;
            bool bAnalogChangeLogChk = AnalogChangeLogDataCheck(sProjectName);
            sw.Stop();
            PrintStep("Check", "Analog Change Log Data", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());

            driver.Dispose();
            

            #region Result judgement
            int iTotalSeleniumAction = dataGridView1.Rows.Count;
            for (int i = 0; i < iTotalSeleniumAction - 1; i++)
            {
                DataGridViewRow row = dataGridView1.Rows[i];
                string sSeleniumResult = row.Cells[2].Value.ToString();
                if (sSeleniumResult != "PASS")
                {
                    bFinalResult = false;
                    EventLog.AddLog("Test Fail !!");
                    EventLog.AddLog("Fail TestItem = " + row.Cells[0].Value.ToString());
                    EventLog.AddLog("BrowserAction = " + row.Cells[1].Value.ToString());
                    EventLog.AddLog("Result = " + row.Cells[2].Value.ToString());
                    EventLog.AddLog("ErrorCode = " + row.Cells[3].Value.ToString());
                    EventLog.AddLog("ExeTime(ms) = " + row.Cells[4].Value.ToString());
                    break;
                }
            }

            if (bFinalResult)
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

        private void Start_Click(object sender, EventArgs e)
        {
            EventLog.AddLog(string.Format("==={0} test start===", sTestItemName));
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog(string.Format("==={0} test end===", sTestItemName));
        }

        private void PrintStep(string sTestItem, string sDescription, bool bResult, string sErrorCode, string sExTime)
        {
            EventLog.AddLog(string.Format("UI Result: {0},{1},{2},{3},{4}", sTestItem, sDescription, bResult, sErrorCode, sExTime));

            DataGridViewRow dgvRow;
            DataGridViewCell dgvCell;

            dgvRow = new DataGridViewRow();

            if (bResult == false)
                dgvRow.DefaultCellStyle.ForeColor = Color.Red;

            dgvCell = new DataGridViewTextBoxCell(); //Column Time

            dgvCell.Value = sTestItem;
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            dgvCell.Value = sDescription;
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            if (bResult)
                dgvCell.Value = "PASS";
            else
                dgvCell.Value = "FAIL";
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            dgvCell.Value = sErrorCode;
            dgvRow.Cells.Add(dgvCell);
            //
            dgvCell = new DataGridViewTextBoxCell();
            dgvCell.Value = sExTime;
            dgvRow.Cells.Add(dgvCell);

            m_DataGridViewCtrlAddDataRow(dgvRow);
        }

        private void InitialRequiredInfo(string sFilePath)
        {
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
            tpc.F_GetPrivateProfileString("ProjectName", "Ground PC or Primary PC", "NA", sDefaultProjectName1, 255, sFilePath);
            tpc.F_GetPrivateProfileString("ProjectName", "Cloud PC or Backup PC", "NA", sDefaultProjectName2, 255, sFilePath);
            tpc.F_GetPrivateProfileString("IP", "Ground PC or Primary PC", "NA", sDefaultIP1, 255, sFilePath);
            tpc.F_GetPrivateProfileString("IP", "Cloud PC or Backup PC", "NA", sDefaultIP2, 255, sFilePath);
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

        private bool AnalogChangeLogDataCheck(string sProjectName)
        {
            bool bCheckData = true;
            string[] ToBeTestTag = { "AT_AI0011", "AT_AO0011", "OPCDA_0011", "OPCUA_0011", "Acc_0011", "ConAna_0011", "SystemSec_0011" };
            //string[] ToBeTestTag = { "SystemSec_0011"};

            for (int i = 0; i < ToBeTestTag.Length; i++)
            {
                EventLog.AddLog("Go to setting page");
                Thread.Sleep(1000);
                driver.FindElement(By.CssSelector("a[href*='/broadWeb/syslog/LogPg.asp?pos=anachglog&ms=1']")).Click();
                Thread.Sleep(1000);
                // select project name
                EventLog.AddLog("select project name");
                new SelectElement(driver.FindElement(By.Name("ProjNameSel"))).SelectByText(sProjectName);
                Thread.Sleep(3000);

                // set today as start date
                string sToday = DateTime.Now.ToString("%d");
                driver.FindElement(By.Name("DateStart")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.LinkText(sToday)).Click(); //click day
                Thread.Sleep(1000);
                EventLog.AddLog("select start date to today: " + sToday);

                //if (ToBeTestTag[i] == "ConDis_0007")
                //{
                //     set start/end time   // 由於離散點是資料有變化則會記錄一次 資料量很大 故設定時間為現在時間往前1分鐘
                //    string sTimeEnd = DateTime.Now.ToString("HH:mm:ss");
                //    string sTimeStart = DateTime.Now.AddMinutes(-1).ToString("HH:mm:ss");
                //    api.ByName("TimeStart").Clear();
                //    api.ByName("TimeStart").Enter(sTimeStart).Exe(); //HHmmss
                //    api.ByName("TimeEnd").Clear();
                //    api.ByName("TimeEnd").Enter(sTimeEnd).Exe();
                //}

                // select one tag to get ODBC data
                EventLog.AddLog("select " + ToBeTestTag[i] + " to get ODBC data");
                driver.FindElement(By.Id("alltags")).Click();
                new SelectElement(driver.FindElement(By.Id("TagNameSel"))).SelectByText(ToBeTestTag[i]);
                driver.FindElement(By.Id("addtag")).Click();
                
                Thread.Sleep(1000);
                driver.FindElement(By.Name("submit")).Click();
                //PrintStep("Set and get Analog Change Log data");
                EventLog.AddLog("Get " + ToBeTestTag[i] + " Analog Change Log data");
                //PrintStep("Login", "login Project Manager page", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());
                Thread.Sleep(10000); // wait to get ODBC data

                driver.FindElement(By.XPath("//*[@id=\"myTable\"]/thead[1]/tr/th[2]/a")).Click();    // click tagname to sort data
                Thread.Sleep(5000);

                bool bRes = bCheckRecordData(ToBeTestTag[i]);
                if (bRes == false)
                    bCheckData = false;

                // print screen
                EventLog.PrintScreen(ToBeTestTag[i] + "_AnalogChangeLog_ODBCData");

                driver.FindElement(By.CssSelector("a[onclick*='../bwproj.asp']")).Click();      //return to homepage
            }

            return bCheckData;
        }

        private bool bCheckRecordData(string sTagName)
        {
            bool bChkTagName = true;
            bool bChkValue = true;

            string sRecordTagNameBefore = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[3]")).Text;
            string sRecordTagName = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[3]")).Text;
            string sRecordTagNameAfter = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[4]/td[3]")).Text;
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
                string sRecordValue1 = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[1]/td[5]")).Text;
                string sRecordValue2 = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[2]/td[5]")).Text;
                string sRecordValue3 = driver.FindElement(By.XPath("//*[@id=\"myTable\"]/tbody/tr[3]/td[5]")).Text;
                EventLog.AddLog(sTagName + " Analog Change log 1: " + sRecordValue1);
                EventLog.AddLog(sTagName + " Analog Change log 2: " + sRecordValue2);
                EventLog.AddLog(sTagName + " Analog Change log 3: " + sRecordValue3);

                if (IsNumeric(sRecordValue1) && IsNumeric(sRecordValue2) && IsNumeric(sRecordValue3))
                {

                    if (sTagName == "SystemSec_0011")
                    {
                        if (((Double.Parse(sRecordValue1) >= 0) && (Double.Parse(sRecordValue1) <= 59)) &&
                            ((Double.Parse(sRecordValue2) >= 0) && (Double.Parse(sRecordValue2) <= 59)) &&
                            ((Double.Parse(sRecordValue3) >= 0) && (Double.Parse(sRecordValue3) <= 59)))
                        {
                            EventLog.AddLog(sTagName + " Record value interval check PASS!!");
                        }
                        else
                        {
                            bChkValue = false;
                            EventLog.AddLog(sTagName + " Record value interval check FAIL!!");
                        }
                    }
                    else
                    {
                        if ((Double.Parse(sRecordValue1) < Double.Parse(sRecordValue2)) && 
                            (Double.Parse(sRecordValue2) < Double.Parse(sRecordValue3)))
                        {
                            EventLog.AddLog(sTagName + " Record value interval check PASS!!");
                        }
                        else
                        {
                            bChkValue = false;
                            EventLog.AddLog(sTagName + " Record value interval check FAIL!!");
                        }
                    }
                }
                else
                {
                    bChkTagName = false;
                    EventLog.AddLog(sTagName + " Record value no Numeric check FAIL!!");
                }
                
            }

            return bChkTagName && bChkValue;
        }

        private bool IsNumeric(string s)
        {
            float output;
            return float.TryParse(s, out output);
        }
    }
}
