using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
//using AdvWebUIAPI;
using ThirdPartyToolControl;
using iATester;
using CommonFunction;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System.Diagnostics;

namespace DownloadSCADA
{
    public partial class Form1 : Form, iATester.iCom
    {
        //IAdvSeleniumAPI api;
        private IWebDriver driver;
        cThirdPartyToolControl tpc = new cThirdPartyToolControl();
        cWACommonFunction wacf = new cWACommonFunction();
        cEventLog EventLog = new cEventLog();

        private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
        private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
        internal const int Max_Rows_Val = 65535;
        string baseUrl;
        string sIniFilePath = @"C:\WebAccessAutoTestSetting.ini";
        bool bPartResult = true;
        bool bFinalResult = true;
        Stopwatch sw = new Stopwatch();
        string sTestItemName = "Download SCADA";
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
            //baseUrl = "http://" + sWebAccessIP;

            //if (sBrowser == "Internet Explorer")
            //{
            //    EventLog.AddLog("Browser= Internet Explorer");
            //    api = new AdvSeleniumAPI("IE", "");
            //    System.Threading.Thread.Sleep(1000);
            //}
            //else if (sBrowser == "Mozilla FireFox")
            //{
            //    EventLog.AddLog("Browser= Mozilla FireFox");
            //    api = new AdvSeleniumAPI("FireFox", "");
            //    System.Threading.Thread.Sleep(1000);
            //}

            //// Launch Firefox and login
            //api.LinkWebUI(baseUrl + "/broadWeb/bwconfig.asp?username=admin");
            //api.ById("userField").Enter("").Submit().Exe();
            //PrintStep("Login WebAccess");

            //api.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
            //PrintStep("Configure project");

            baseUrl = "http://" + sWebAccessIP;
            //baseUrl = "http://172.16.12.11" ;

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
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(60)); // Set implicit wait timeouts to 5 secs
                //driver.Manage().Timeouts().SetScriptTimeout(new TimeSpan(0, 0, 0, 5));  // Set script timeouts to 5 secs

                //WebDriverWait _wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                //string text = _wait.Until(d => d.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwconfig.asp?username=admin')]"))).Text;
                //_wait.Until(baseUrl + "/broadWeb/bwRoot.asp?username=admin");

                driver.Navigate().GoToUrl(baseUrl + "/broadWeb/bwRoot.asp?username=admin");
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwconfig.asp?username=admin')]")).Click();
                Thread.Sleep(1000);
                driver.FindElement(By.Id("userField")).Submit();
                Thread.Sleep(1000);
                driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwMain.asp') and contains(@href, 'ProjName=" + sProjectName + "')]")).Click();
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

            sw.Reset(); sw.Start(); bPartResult = true;
            try
            {
                EventLog.AddLog("Start Download...");
                bPartResult = wacf.Download(driver, slanguage);
                //bPartResult = wacf.StopKernel(driver);
                //bPartResult = wacf.StartKernel(driver);
            }
            catch (Exception ex)
            {
                EventLog.AddLog(@"Error occurred Downloading : " + ex.ToString());
                bPartResult = false;
            }
            sw.Stop();
            PrintStep("Download", "Download Project", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());

            //api.Quit();
            //PrintStep("Quit browser");
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

        private void ReturnSCADAPage()
        {
            //api.SwitchToCurWindow(0);
            //api.SwitchToFrame("leftFrame", 0);
            //api.ByXpath("//a[contains(@href, '/broadWeb/bwMainRight.asp') and contains(@href, 'name=TestSCADA')]").Click();

        }

        private void Start_Click(object sender, EventArgs e)
        {
            long lErrorCode = 0;
            EventLog.AddLog("===Download SCADA start===");
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===Download SCADA end===");
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
