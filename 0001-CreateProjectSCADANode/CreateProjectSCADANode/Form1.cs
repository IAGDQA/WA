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
using ThirdPartyToolControl;
using iATester;
using System.Runtime.InteropServices;

namespace CreateProjectSCADANode
{
    public partial class Form1 : Form, iATester.iCom
    {
        IAdvSeleniumAPI api;
        cThirdPartyToolControl tpc = new cThirdPartyToolControl();
        private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
        private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
        internal const int Max_Rows_Val = 65535;
        string baseUrl;
        string sIniFilePath = @"C:\WebAccessAutoTestSetting.ini";
        string slanguage;

        [DllImport("kernel32")]
        public static extern uint GetTickCount();

        //Send Log data to iAtester
        public event EventHandler<LogEventArgs> eLog = delegate { };
        //Send test result to iAtester
        public event EventHandler<ResultEventArgs> eResult = delegate { };
        //Send execution status to iAtester
        public event EventHandler<StatusEventArgs> eStatus = delegate { };

        public void StartTest()
        {
            //Add test code
            long lErrorCode = (long)ErrorCode.SUCCESS;
            EventLog.AddLog("===Create Project and SCADA node start (by iATester)===");
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text, UserEmail.Text);
            EventLog.AddLog("===Create Project and SCADA node end (by iATester)===");

            if (lErrorCode == 0)
            {
                eResult(this, new ResultEventArgs(iResult.Pass));
                eStatus(this, new StatusEventArgs(iStatus.Completion));
            }
            else
            {
                eResult(this, new ResultEventArgs(iResult.Fail));
                eStatus(this, new StatusEventArgs(iStatus.Stop));
            }

            //eStatus(this, new StatusEventArgs(iStatus.Completion));
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
            comboBox_language.SelectedIndex = 0;

            if (System.IO.File.Exists(sIniFilePath))
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
        }

        long Form1_Load(string sProjectName, string sWebAccessIP, string sTestLogFolder, string sBrowser, string sUserEmail)
        {
            baseUrl = "http://" + sWebAccessIP;
            if (sBrowser == "Internet Explorer")
            {
                EventLog.AddLog("Browser= Internet Explorer");
                //driver = new FirefoxDriver();
                api = new AdvSeleniumAPI("IE", "");
                System.Threading.Thread.Sleep(1000);
            }
            else if (sBrowser == "Mozilla FireFox")
            {
                EventLog.AddLog("Browser= Mozilla FireFox");
                //driver = new FirefoxDriver();
                api = new AdvSeleniumAPI("FireFox", "");
                System.Threading.Thread.Sleep(1000);
            }

            // Launch Firefox and login
            api.LinkWebUI(baseUrl + "/broadWeb/bwconfig.asp?username=admin");
            api.ById("userField").Enter("").Submit().Exe();
            PrintStep("Login WebAccess");

            //Step0: LogData Maintenance setting
            EventLog.AddLog("Go to LogData Maintenance page");
            SetLogDataMaintenance(sTestLogFolder);
            PrintStep("Set LogData Maintenance");

            //Step1
            EventLog.AddLog("Create Project Node...");
            CreateProject(sProjectName, sWebAccessIP);
            PrintStep("Create Project Node");
            Thread.Sleep(1000);
            
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
            PrintStep("Configure project");
            Thread.Sleep(500);

            //Step2
            EventLog.AddLog("Create SCADA Node...");
            CreateSCADANode(sWebAccessIP, sUserEmail);

            /* Because of frequent timeout issue of creating SCADA node, use the mechanism judgement instead of checking selenium result */
            //Step3 check if scada node exist
            bool bResult = ReturnSCADAPage(20000);
            PrintStep("CheckSCADANode");

            api.Quit();
            //PrintStep("Quit browser");

            if (bResult)
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

        private void SetLogDataMaintenance(string sTestLogFolder)
        {
            api.ByXpath("//a[contains(@href, '/broadWeb/syslog/ArchPg.asp')]").Click();

            //DataLog Trend
            api.ByName("TrendArchChkAll").Click();
            Thread.Sleep(500);
            api.ByName("ArchFolder1").Clear();
            api.ByName("ArchFolder1").Enter(sTestLogFolder).Exe();

            //ODBC Log
            api.ByName("ODBCArchChkAll").Click();
            Thread.Sleep(500);
            api.ByName("ODBCDeleChkAll").Click();
            Thread.Sleep(500);
            api.ByName("ODBCExpTimeType1").Click();
            api.ByName("ODBCExpTimeType2").Click();
            api.ByName("ODBCExpTimeType3").Click();
            api.ByName("ODBCExpTimeType4").Click();
            api.ByName("ODBCExpTimeType5").Click();
            api.ByName("ODBCExpTimeType6").Click();
            api.ByName("ODBCExpTimeType7").Click();
            api.ByName("ODBCExpTimeType8").Click();
            api.ByName("ArchFolder2").Clear();
            api.ByName("ArchFolder2").Enter(sTestLogFolder).Exe();

            //Excel Report Maintenance
            api.ByName("ExcelArchChkAll").Click();
            Thread.Sleep(500);
            api.ByName("ArchFolder3").Clear();
            api.ByName("ArchFolder3").Enter(sTestLogFolder).Submit().Exe();
            
        }

        private void CreateProject(string sProjectName, string sWebAccessIP)
        {
            EventLog.AddLog("Create a new project");
            // Create a new project
            //driver.FindElement(By.Name("ProjName")).Clear();
            //driver.FindElement(By.Name("ProjName")).SendKeys(projName);
            //driver.FindElement(By.Name("ProjIPLong")).Clear();
            //driver.FindElement(By.Name("ProjIPLong")).SendKeys(sWebAccessIP);
            //driver.FindElement(By.Name("ProjIPLong")).Submit();
            api.ByName("ProjName").Clear();
            int i = api.ByName("ProjName").Enter(sProjectName).Exe();
            //PrintStep("123123");
            api.ByName("ProjIPLong").Clear();
            api.ByName("ProjIPLong").Enter(sWebAccessIP).Submit().Exe();
            //api.ByName("ProjIPLong").Submit().Exe();

            // Confirm to create
            //string alertText = driver.SwitchTo().Alert().Text;
            string alertText = api.GetAlartTxt();

            switch (slanguage)
            {
                case "ENG":
                    if (alertText == "Do you want to create a new project ( Project Name : " + sProjectName + " )? ")
                        api.Accept();
                    break;
                case "CHT":
                    if (alertText == "您要建立新的工程 ( 工程名稱 : " + sProjectName + " )? ")
                        api.Accept();
                    break;
                case "CHS":
                    if (alertText == "你想建立新的工程 ( 工程名称 : " + sProjectName + " )? ")
                        api.Accept();
                    break;
                case "JPN":
                    if (alertText == "新しいﾌﾟﾛｼﾞｪｸﾄを作成しますか ( ﾌﾟﾛｼﾞｪｸﾄ名 : " + sProjectName + " )? ")
                        api.Accept();
                    break;
                case "KRN":
                    if (alertText == "새 프로젝트를 생성할까요? ( 프로젝트명 : " + sProjectName + " )? ")
                        api.Accept();
                    break;
                case "FRN":

                default:
                    if (alertText == "Do you want to create a new project ( Project Name : " + sProjectName + " )? ")
                        api.Accept();
                    break;
            }
        }

        private void CreateSCADANode(string sWebAccessIP, string sUserEmail)
        {
            // Create SCADA node
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//a[2]/font/b").Click();
            Thread.Sleep(1000);
            api.ByName("NodeName").Clear();
            api.ByName("NodeName").Enter("TestSCADA").Exe();
            api.ByName("AddressLong").Clear();
            api.ByName("AddressLong").Enter(sWebAccessIP).Exe();

            api.ByName("EMAIL_SERVER").Clear();                         //Outgoing Email (SMTP) Server
            api.ByName("EMAIL_SERVER").Enter("smtp.mail.yahoo.com").Exe();
            api.ByName("EMAIL_PORT").Clear();
            api.ByName("EMAIL_PORT").Enter("587").Exe();
            api.ByName("EMAIL_ADDRESS").Clear();
            api.ByName("EMAIL_ADDRESS").Enter("webaccess2016@yahoo.com").Exe();
            api.ByName("EMAIL_USER").Clear();                         //Email Account Name
            api.ByName("EMAIL_USER").Enter("webaccess2016").Exe();

            //api.ByXpath("input[@name='D_EMAIL_PASSWORD']").Clear();   // Email Password   // 目前有問題無法成功輸入
            //api.ByXpath("input[@name='D_EMAIL_PASSWORD']").Enter("123214").Exe();
            //api.ByName("D_MAIL_PASSWORD").Clear();
            //api.ByName("D_MAIL_PASSWORD").Enter("xxxx").Exe();
            //api.ByName("D_MAIL_PASSWORDB").Clear();
            //api.ByName("D_MAIL_PASSWORDB").Enter("xxxx").Exe();
            api.ByName("EMAIL_FROM").Clear();
            api.ByName("EMAIL_FROM").Enter("webaccess2016@yahoo.com").Exe();
            api.ByName("EMAIL_TO_SRPT").Clear();
            api.ByName("EMAIL_TO_SRPT").Enter("webaccess2016@yahoo.com").Exe();
            api.ByName("EMAIL_TO").Clear();
            api.ByName("EMAIL_TO").Enter("webaccess2016@yahoo.com").Exe();
            api.ByName("EMAIL_CC").Clear();
            api.ByName("EMAIL_CC").Enter(sUserEmail).Exe();

            api.ByName("ALARM_LOG_TO_ODBC").Click();
            api.ByName("CHANGE_LOG_TO_ODBC").Click();
            api.ByName("DATA_LOG_TO_ODBC").Click();
            api.ByName("DATA_LOG_USE_RTDB").Click();
            EventLog.AddLog("Enter some required info done and submit start");
            int iSubmitResult = api.ByName("Max_Number_of_Client").Enter("").Submit().Exe();    // for IE special
            if (iSubmitResult == 0)
                EventLog.AddLog("submit success!!");
            else
                EventLog.AddLog("submit fail!!");

            Thread.Sleep(1000);
            PrintStep("Create SCADA Node");
        }

        private void Start_Click(object sender, EventArgs e)
        {
            long lErrorCode = (long)ErrorCode.SUCCESS;
            EventLog.AddLog("===Create Project and SCADA node start===");
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text, UserEmail.Text);
            EventLog.AddLog("===Create Project and SCADA node end===");
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

        private void InitialRequiredInfo(string sFilePath)
        {
            StringBuilder sDefaultUserLanguage = new StringBuilder(255);
            StringBuilder sDefaultUserEmail = new StringBuilder(255);
            StringBuilder sDefaultProjectName1 = new StringBuilder(255);
            StringBuilder sDefaultProjectName2 = new StringBuilder(255);
            StringBuilder sDefaultIP1 = new StringBuilder(255);
            StringBuilder sDefaultIP2 = new StringBuilder(255);
            StringBuilder sDefaultIP3 = new StringBuilder(255);
            /*
            tpc.F_WritePrivateProfileString("ProjectName", "Ground PC or Primary PC", "TestProject", @"C:\WebAccessAutoTestSetting.ini");
            tpc.F_WritePrivateProfileString("ProjectName", "Cloud PC or Backup PC", "CTestProject", @"C:\WebAccessAutoTestSetting.ini");
            tpc.F_WritePrivateProfileString("IP", "Ground PC or Primary PC", "172.18.3.62", @"C:\WebAccessAutoTestSetting.ini");
            tpc.F_WritePrivateProfileString("IP", "Cloud PC or Backup PC", "172.18.3.65", @"C:\WebAccessAutoTestSetting.ini");
            */
            tpc.F_GetPrivateProfileString("UserInfo", "Language", "NA", sDefaultUserLanguage, 255, sFilePath);
            tpc.F_GetPrivateProfileString("UserInfo", "Email", "NA", sDefaultUserEmail, 255, sFilePath);
            tpc.F_GetPrivateProfileString("ProjectName", "Ground PC or Primary PC", "NA", sDefaultProjectName1, 255, sFilePath);
            tpc.F_GetPrivateProfileString("IP", "Ground PC or Primary PC", "NA", sDefaultIP1, 255, sFilePath);
            tpc.F_GetPrivateProfileString("IP", "Cloud PC or Backup PC", "NA", sDefaultIP2, 255, sFilePath);
            tpc.F_GetPrivateProfileString("IP", "Redundant Secondary PC", "NA", sDefaultIP3, 255, sFilePath);
            slanguage = sDefaultUserLanguage.ToString();    // 在這邊讀取使用語言
            UserEmail.Text = sDefaultUserEmail.ToString();
            ProjectName.Text = sDefaultProjectName1.ToString();
            WebAccessIP.Text = sDefaultIP1.ToString();
            textBox_CloudPC_IP.Text = sDefaultIP2.ToString();
            textBox_BackupPC_IP.Text = sDefaultIP3.ToString();
        }

        private bool ReturnSCADAPage(int iTimeout)
        {
            long lStartTestTime = GetTickCount();
            long lEndTestTime;
            int iCheckIfSCADAExis = -1;
            do
            {
                EventLog.AddLog("Try...");
                api.SwitchToCurWindow(0);
                api.SwitchToFrame("leftFrame", 0);
                iCheckIfSCADAExis = api.ByXpath("//a[contains(@href, '/broadWeb/bwMainRight.asp') and contains(@href, 'name=TestSCADA')]").Click();
                lEndTestTime = GetTickCount();
            } while (iCheckIfSCADAExis != 0 && (lEndTestTime-lStartTestTime) < iTimeout);

            if (iCheckIfSCADAExis == 0)
                return true;
            else
                return false;
        }

        private void CheckifIniFileChange()
        {
            StringBuilder sDefaultUserLanguage = new StringBuilder(255);
            StringBuilder sDefaultUserEmail = new StringBuilder(255);
            StringBuilder sDefaultProjectName1 = new StringBuilder(255);
            StringBuilder sDefaultProjectName2 = new StringBuilder(255);
            StringBuilder sDefaultIP1 = new StringBuilder(255);
            StringBuilder sDefaultIP2 = new StringBuilder(255); // cloud pc ip
            StringBuilder sDefaultIP3 = new StringBuilder(255); // backup pc ip

            if (System.IO.File.Exists(sIniFilePath))    // 比對ini檔與ui上的值是否相同
            {
                EventLog.AddLog(".ini file exist, check if .ini file need to update");
                tpc.F_GetPrivateProfileString("UserInfo", "Language", "NA", sDefaultUserLanguage, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("UserInfo", "Email", "NA", sDefaultUserEmail, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("ProjectName", "Ground PC or Primary PC", "NA", sDefaultProjectName1, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("ProjectName", "Cloud PC or Backup PC", "NA", sDefaultProjectName2, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("IP", "Ground PC or Primary PC", "NA", sDefaultIP1, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("IP", "Cloud PC or Backup PC", "NA", sDefaultIP2, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("IP", "Redundant Secondary PC", "NA", sDefaultIP3, 255, sIniFilePath);

                if (comboBox_language.Text != sDefaultUserLanguage.ToString())
                {
                    tpc.F_WritePrivateProfileString("UserInfo", "Language", comboBox_language.Text, sIniFilePath);
                    EventLog.AddLog("New Language update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultUserLanguage.ToString());
                    EventLog.AddLog("New ini:" + comboBox_language.Text);
                }
                if (UserEmail.Text != sDefaultUserEmail.ToString())
                {
                    tpc.F_WritePrivateProfileString("UserInfo", "Email", UserEmail.Text, sIniFilePath);
                    EventLog.AddLog("New UserEmail update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultUserEmail.ToString());
                    EventLog.AddLog("New ini:" + UserEmail.Text);
                }
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
                if (textBox_CloudPC_IP.Text != sDefaultIP2.ToString())
                {
                    tpc.F_WritePrivateProfileString("IP", "Cloud PC or Backup PC", textBox_CloudPC_IP.Text, sIniFilePath);
                    EventLog.AddLog("New WebAccess Cloud PC IP update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultIP2.ToString());
                    EventLog.AddLog("New ini:" + textBox_CloudPC_IP.Text);
                }
                if (textBox_BackupPC_IP.Text != sDefaultIP3.ToString())
                {
                    tpc.F_WritePrivateProfileString("IP", "Redundant Secondary PC", textBox_BackupPC_IP.Text, sIniFilePath);
                    EventLog.AddLog("New WebAccess Backup PC IP update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultIP3.ToString());
                    EventLog.AddLog("New ini:" + textBox_BackupPC_IP.Text);
                }
            }
            else
            {   // 若ini檔不存在 則建立新的
                EventLog.AddLog(".ini file not exist, create new .ini file. Path: " + sIniFilePath);
                tpc.F_WritePrivateProfileString("UserInfo", "Language", comboBox_language.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("UserInfo", "Email", UserEmail.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("ProjectName", "Ground PC or Primary PC", ProjectName.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("ProjectName", "Cloud PC or Backup PC", "CTestProject", sIniFilePath);
                tpc.F_WritePrivateProfileString("ProjectName", "Redundant Secondary PC", "TestProject_bk", sIniFilePath);
                tpc.F_WritePrivateProfileString("IP", "Ground PC or Primary PC", WebAccessIP.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("IP", "Cloud PC or Backup PC", textBox_CloudPC_IP.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("IP", "Redundant Secondary PC", textBox_BackupPC_IP.Text, sIniFilePath);
            }
        }
    }
}
