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

namespace CreateExcelReport
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
            EventLog.AddLog("===Create Excel Report start (by iATester)===");
            if (System.IO.File.Exists(sIniFilePath))    // 再load一次
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text, UserEmail.Text);
            EventLog.AddLog("===Create Excel Report end (by iATester)===");

            Thread.Sleep(3000);

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

            // Configure project by project name
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
            PrintStep("Configure project");

            //Create Excel Report
            
            //step 1
            EventLog.AddLog("Create DailyReport Excel Report...");
            Create_DailyReport_ExcelReport(sUserEmail);
            PrintStep("Create DailyReport Excel Report");

            ReturnSCADAPage();
            
            //step 2
            EventLog.AddLog("Create SelfDefined Excel Report...");
            Create_SelfDefined_ExcelReport(sUserEmail);
            PrintStep("Create SelfDefined Excel Report");
            PrintScreen("ExcelReportTest1", sTestLogFolder);
            /*
            api.ByXpath("//a[contains(text(),'Generate')]").Click();
            Thread.Sleep(500);

            PrintScreen("ExcelReportTest2", sTestLogFolder);

            // copy ExcelReport_SelfDefined_DataLog_Excel_ex.xlsx to log path.
            //C:\inetpub\wwwroot\broadweb\WaExlViewer\report\TestProject_TestSCADA\ExcelReport_SelfDefined_DataLog_Excel
            {
                string fileNameSrc = "ExcelReport_SelfDefined_DataLog_Excel_ex.xlsx";
                string fileNameTar = string.Format("ExcelReport_SelfDefined_DataLog_Excel_ex_{0:yyyyMMdd_hhmmss}.xlsx", DateTime.Now);

                string sourcePath = string.Format(@"C:\inetpub\wwwroot\broadweb\WaExlViewer\report\{0}_TestSCADA\ExcelReport_SelfDefined_DataLog_Excel", sProjectName);
                string targetPath = sTestLogFolder;

                // Use Path class to manipulate file and directory paths.
                string sourceFile = System.IO.Path.Combine(sourcePath, fileNameSrc);
                string destFile = System.IO.Path.Combine(targetPath, fileNameTar);
                System.IO.File.Copy(sourceFile, destFile, true);
            }
            */
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

            if (bSeleniumResult)
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

        private void PrintScreen(string sFileName, string sFilePath)
        {
            Bitmap myImage = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
            Graphics g = Graphics.FromImage(myImage);
            g.CopyFromScreen(new Point(0, 0), new Point(0, 0), new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height));
            IntPtr dc1 = g.GetHdc();
            g.ReleaseHdc(dc1);
            //myImage.Save(@"c:\screen0.jpg");
            myImage.Save(string.Format("{0}\\{1}_{2:yyyyMMdd_hhmmss}.jpg", sFilePath, sFileName, DateTime.Now));
        }

        private void Create_DailyReport_ExcelReport(string sUserEmail)
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/WaExlViewer/WaExlViewer.asp')]").Click();
            
            for(int t = 1; t<=4; t++)   /////// template 1~4 ; 8min
            {
                api.ByXpath("//a[contains(@href, 'addRptCfg.aspx')]").Click();

                api.ByName("rptName").Clear();
                api.ByName("rptName").Enter(string.Format("ER_ODBC_T{0}_DR_8min_LastValue_ER", t)).Exe();
                api.ByXpath("(//input[@name='dataSrc'])[2]").Click();   // ODBC
                api.ByName("selectTemplate").SelectVal(string.Format("template{0}.xlsx", t)).Exe();

                switch (slanguage)
                {
                    case "ENG":
                        api.ByName("speTimeFmt").SelectTxt("Daily Report").Exe();
                        break;
                    case "CHT":
                        api.ByName("speTimeFmt").SelectTxt("日報表").Exe();
                        break;
                    case "CHS":
                        api.ByName("speTimeFmt").SelectTxt("日报表").Exe();
                        break;
                    case "JPN":
                        api.ByName("speTimeFmt").SelectTxt("日報").Exe();
                        break;
                    case "KRN":
                        api.ByName("speTimeFmt").SelectTxt("일간 보고").Exe();
                        break;
                    case "FRN":

                    default:
                        api.ByName("speTimeFmt").SelectTxt("Daily Report").Exe();
                        break;
                }

                // set end date
                string sToday = string.Format("{0:dd}", DateTime.Now);
                int iToday = Int32.Parse(sToday);   // 為了讓讀出來的日期去掉第一個零 ex: "06" -> "6"
                string ssToday = string.Format("{0}", iToday);
                api.ByName("tEnd").Click();
                Thread.Sleep(500);
                api.ByXpath("//button[@type='button']").Click();
                api.ByXpath("(//button[@type='button'])[2]").Click();

                api.ByName("interval").Clear();
                api.ByName("interval").Enter("8").Exe();                // Time Interval = 8
                api.ByXpath("(//input[@name='fileType'])[2]").Click();  // Time Unit: Minute
                api.ByXpath("//input[@name='valueType']").Click();      // Data Type: Last Value
                // 湊到最大上限32個TAG
                string[] ReportTagName = { "Calc_ConAna", "Calc_ConDis", "Calc_ModBusAI", "Calc_ModBusAO", "Calc_ModBusDI", "Calc_ModBusDO", "Calc_OPCDA", "Calc_OPCUA","Calc_System",
                                         "Calc_ANDConDisAll", "Calc_ANDDIAll", "Calc_ANDDOAll", "Calc_AvgAIAll", "Calc_AvgAOAll", "Calc_AvgConAnaAll", "Calc_AvgSystemAll",
                                         "AT_AI0050", "AT_AI0100", "AT_AI0150", "AT_AO0050", "AT_AO0100", "AT_AO0150", "AT_DI0050", "AT_DI0100", "AT_DI0150", "AT_DO0050",
                                         "AT_DO0100", "AT_DO0150", "ConAna_0100", "ConDis_0100", "SystemSec_0100", "SystemSec_0200"};

                for (int i = 0; i < ReportTagName.Length; i++)
                {
                    api.ById("tagsLeftList").SelectTxt(ReportTagName[i]).Exe();
                }
                api.ById("torightBtn").Click();

                api.ById("tagsRightList").SelectTxt(ReportTagName[0]).Exe();

                api.ByXpath("(//input[@name='attachFormat'])[2]").Click();  // Send Email: Excel Report
                api.ByName("emailto").Clear();
                api.ByName("emailto").Enter(sUserEmail).Exe();
                //api.ByName("rptName").Enter("").Submit().Exe();   這邊不能用這種方式submit...很奇怪 反而按submit的button可以正常送出
                api.ByName("cfgSubmit").Click();
            }

            for (int t = 1; t <= 4; t++)   ////// template 1~4 ; 1hour
            {
                api.ByXpath("//a[contains(@href, 'addRptCfg.aspx')]").Click();

                api.ByName("rptName").Clear();
                api.ByName("rptName").Enter(string.Format("ER_ODBC_T{0}_DR_1hour_LastValue_ER", t)).Exe();
                api.ByXpath("(//input[@name='dataSrc'])[2]").Click();   // ODBC
                api.ByName("selectTemplate").SelectVal(string.Format("template{0}.xlsx", t)).Exe();
                switch (slanguage)
                {
                    case "ENG":
                        api.ByName("speTimeFmt").SelectTxt("Daily Report").Exe();
                        break;
                    case "CHT":
                        api.ByName("speTimeFmt").SelectTxt("日報表").Exe();
                        break;
                    case "CHS":
                        api.ByName("speTimeFmt").SelectTxt("日报表").Exe();
                        break;
                    case "JPN":
                        api.ByName("speTimeFmt").SelectTxt("日報").Exe();
                        break;
                    case "KRN":
                        api.ByName("speTimeFmt").SelectTxt("일간 보고").Exe();
                        break;
                    case "FRN":

                    default:
                        api.ByName("speTimeFmt").SelectTxt("Daily Report").Exe();
                        break;
                }

                // set end date
                string sToday = string.Format("{0:dd}", DateTime.Now);
                int iToday = Int32.Parse(sToday);   // 為了讓讀出來的日期去掉第一個零 ex: "06" -> "6"
                string ssToday = string.Format("{0}", iToday);
                api.ByName("tEnd").Click();
                Thread.Sleep(500);
                api.ByXpath("//button[@type='button']").Click();
                api.ByXpath("(//button[@type='button'])[2]").Click();

                api.ByName("interval").Clear();
                api.ByName("interval").Enter("1").Exe();                // Time Interval = 1
                api.ByXpath("(//input[@name='fileType'])[3]").Click();  // Time Unit: hour
                api.ByXpath("//input[@name='valueType']").Click();      // Data Type: Last Value
                // 湊到最大上限32個TAG
                string[] ReportTagName = { "Calc_ConAna", "Calc_ConDis", "Calc_ModBusAI", "Calc_ModBusAO", "Calc_ModBusDI", "Calc_ModBusDO", "Calc_OPCDA", "Calc_OPCUA","Calc_System",
                                         "Calc_ANDConDisAll", "Calc_ANDDIAll", "Calc_ANDDOAll", "Calc_AvgAIAll", "Calc_AvgAOAll", "Calc_AvgConAnaAll", "Calc_AvgSystemAll",
                                         "AT_AI0050", "AT_AI0100", "AT_AI0150", "AT_AO0050", "AT_AO0100", "AT_AO0150", "AT_DI0050", "AT_DI0100", "AT_DI0150", "AT_DO0050",
                                         "AT_DO0100", "AT_DO0150", "ConAna_0100", "ConDis_0100", "SystemSec_0100", "SystemSec_0200"};

                for (int i = 0; i < ReportTagName.Length; i++)
                {
                    try
                    {
                        api.ById("tagsLeftList").SelectTxt(ReportTagName[i]).Exe();
                    }
                    catch (Exception ex)
                    {
                        i--;
                    }
                }
                api.ById("torightBtn").Click();
                api.ById("tagsRightList").SelectTxt(ReportTagName[0]).Exe();

                api.ByXpath("(//input[@name='attachFormat'])[2]").Click();  // Send Email: Excel Report
                api.ByName("emailto").Clear();
                api.ByName("emailto").Enter(sUserEmail).Exe();
                api.ByName("cfgSubmit").Click();
            }
            
            for (int d = 1; d <= 3; d++)   ////// data type = Maximum or  Minimum or  Average
            {
                api.ByXpath("//a[contains(@href, 'addRptCfg.aspx')]").Click();

                api.ByName("rptName").Clear();

                if(d == 1)
                    api.ByName("rptName").Enter("ER_ODBC_T1_DR_1hour_Max_ER").Exe();
                else if(d == 2)
                    api.ByName("rptName").Enter("ER_ODBC_T1_DR_1hour_Min_ER").Exe();
                else if(d == 3)
                    api.ByName("rptName").Enter("ER_ODBC_T1_DR_1hour_Avg_ER").Exe();

                api.ByXpath("(//input[@name='dataSrc'])[2]").Click();   // ODBC
                api.ByName("selectTemplate").SelectVal("template1.xlsx").Exe();
                switch (slanguage)
                {
                    case "ENG":
                        api.ByName("speTimeFmt").SelectTxt("Daily Report").Exe();
                        break;
                    case "CHT":
                        api.ByName("speTimeFmt").SelectTxt("日報表").Exe();
                        break;
                    case "CHS":
                        api.ByName("speTimeFmt").SelectTxt("日报表").Exe();
                        break;
                    case "JPN":
                        api.ByName("speTimeFmt").SelectTxt("日報").Exe();
                        break;
                    case "KRN":
                        api.ByName("speTimeFmt").SelectTxt("일간 보고").Exe();
                        break;
                    case "FRN":

                    default:
                        api.ByName("speTimeFmt").SelectTxt("Daily Report").Exe();
                        break;
                }

                // set end date
                string sToday = string.Format("{0:dd}", DateTime.Now);
                int iToday = Int32.Parse(sToday);   // 為了讓讀出來的日期去掉第一個零 ex: "06" -> "6"
                string ssToday = string.Format("{0}", iToday);
                api.ByName("tEnd").Click();
                Thread.Sleep(500);
                api.ByXpath("//button[@type='button']").Click();
                api.ByXpath("(//button[@type='button'])[2]").Click();

                api.ByName("interval").Clear();
                api.ByName("interval").Enter("1").Exe();                // Time Interval = 1
                api.ByXpath("(//input[@name='fileType'])[3]").Click();  // Time Unit: hour
                api.ByXpath(string.Format("(//input[@name='valueType'])[{0}]", d+1)).Click();      // Data Type
                // 湊到最大上限32個TAG
                string[] ReportTagName = { "Calc_ConAna", "Calc_ConDis", "Calc_ModBusAI", "Calc_ModBusAO", "Calc_ModBusDI", "Calc_ModBusDO", "Calc_OPCDA", "Calc_OPCUA","Calc_System",
                                         "Calc_ANDConDisAll", "Calc_ANDDIAll", "Calc_ANDDOAll", "Calc_AvgAIAll", "Calc_AvgAOAll", "Calc_AvgConAnaAll", "Calc_AvgSystemAll",
                                         "AT_AI0050", "AT_AI0100", "AT_AI0150", "AT_AO0050", "AT_AO0100", "AT_AO0150", "AT_DI0050", "AT_DI0100", "AT_DI0150", "AT_DO0050",
                                         "AT_DO0100", "AT_DO0150", "ConAna_0100", "ConDis_0100", "SystemSec_0100", "SystemSec_0200"};

                for (int i = 0; i < ReportTagName.Length; i++)
                {
                    try
                    {
                        api.ById("tagsLeftList").SelectTxt(ReportTagName[i]).Exe();
                    }
                    catch (Exception ex)
                    {
                        i--;
                    }
                }
                api.ById("torightBtn").Click();
                api.ById("tagsRightList").SelectTxt(ReportTagName[0]).Exe();

                api.ByXpath("(//input[@name='attachFormat'])[2]").Click();  // Send Email: Excel Report
                api.ByName("emailto").Clear();
                api.ByName("emailto").Enter(sUserEmail).Exe();
                api.ByName("cfgSubmit").Click();
            }
        }

        private void Create_SelfDefined_ExcelReport(string sUserEmail)
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/WaExlViewer/WaExlViewer.asp')]").Click();

            EventLog.AddLog("Create SelfDefined Data log ExcelReport");
            Create_SelfDefined_Datalog_ExcelReport(sUserEmail);
            PrintStep("Create SelfDefined Data log ExcelReport");

            EventLog.AddLog("Create SelfDefined ODBC Data ExcelReport");
            Create_SelfDefined_ODBCData_ExcelReport(sUserEmail);
            PrintStep("Create SelfDefined ODBC Data ExcelReport");

            EventLog.AddLog("Create SelfDefined Alarm ExcelReport");
            Create_SelfDefined_Alarm_ExcelReport(sUserEmail);
            PrintStep("Create SelfDefined Alarm ExcelReport");

            EventLog.AddLog("Create SelfDefined Action Log ExcelReport");
            Create_SelfDefined_ActionLog_ExcelReport(sUserEmail);
            PrintStep("Create SelfDefined Action Log ExcelReport");

            EventLog.AddLog("Create SelfDefined Event Log ExcelReport");
            Create_SelfDefined_EventLog_ExcelReport(sUserEmail);
            PrintStep("Create SelfDefined Event Log ExcelReport");
        }

        private void Create_SelfDefined_Datalog_ExcelReport(string sUserEmail)
        {
            //api.SwitchToCurWindow(0);
            //api.SwitchToFrame("rightFrame", 0);
            //api.ByXpath("//a[contains(@href, '/broadWeb/WaExlViewer/WaExlViewer.asp')]").Click();
            api.ByXpath("//a[contains(@href, 'addRptCfg.aspx')]").Click();

            api.ByName("rptName").Clear();
            api.ByName("rptName").Enter("ExcelReport_SelfDefined_DataLog_Excel").Exe();
            api.ByName("dataSrc").Click();  // Click "Data Log" button
            api.ByName("selectTemplate").SelectVal("template1.xlsx").Exe();
            switch (slanguage)
            {
                case "ENG":
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
                case "CHT":
                    api.ByName("speTimeFmt").SelectTxt("自訂").Exe();
                    break;
                case "CHS":
                    api.ByName("speTimeFmt").SelectTxt("自订").Exe();
                    break;
                case "JPN":
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
                case "KRN":
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
                case "FRN":

                default:
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
            }

            // set today as start/end date
            string sToday = string.Format("{0:dd}", DateTime.Now);
            string sTomorrow = string.Format("{0:dd}", DateTime.Now.AddHours(8).AddDays(+1));
            int iToday = Int32.Parse(sToday);   // 為了讓讀出來的日期去掉第一個零 ex: "06" -> "6"
            int iTomorrow = Int32.Parse(sTomorrow);
            string ssToday = string.Format("{0}", iToday);
            string ssTomorrow = string.Format("{0}", iTomorrow);
            api.ByName("tStart").Click();
            Thread.Sleep(500);
            api.ByTxt(ssToday).Click();
            Thread.Sleep(500);
            api.ByXpath("(//button[@type='button'])[2]").Click();

            if (iTomorrow != 1)
            {
                api.ByName("tEnd").Click();
                Thread.Sleep(500);
                api.ByTxt(ssTomorrow).Click();
                Thread.Sleep(500);
                api.ByXpath("(//button[@type='button'])[2]").Click();
            }
            else  // 跳頁
            {
                api.ByName("tEnd").Click();
                Thread.Sleep(500);
                api.ByCss("span.ui-icon.ui-icon-circle-triangle-e").Click();
                Thread.Sleep(500);
                api.ByTxt(ssTomorrow).Click();
                Thread.Sleep(500);
                api.ByXpath("(//button[@type='button'])[2]").Click();
            }

            api.ByName("interval").Clear();
            api.ByName("interval").Enter("1").Exe();
            api.ByXpath("(//input[@name='fileType'])[2]").Click();

            // 湊到最大上限32個TAG
            string[] ReportTagName = { "Calc_ConAna", "Calc_ConDis", "Calc_ModBusAI", "Calc_ModBusAO", "Calc_ModBusDI", "Calc_ModBusDO", "Calc_OPCDA", "Calc_OPCUA","Calc_System",
                                     "Calc_ANDConDisAll", "Calc_ANDDIAll", "Calc_ANDDOAll", "Calc_AvgAIAll", "Calc_AvgAOAll", "Calc_AvgConAnaAll", "Calc_AvgSystemAll",
                                     "AT_AI0050", "AT_AI0100", "AT_AI0150", "AT_AO0050", "AT_AO0100", "AT_AO0150", "AT_DI0050", "AT_DI0100", "AT_DI0150", "AT_DO0050",
                                     "AT_DO0100", "AT_DO0150", "ConAna_0100", "ConDis_0100", "SystemSec_0100", "SystemSec_0200"};
            for (int i = 0; i < ReportTagName.Length; i++)
            {
                try
                {
                    api.ById("tagsLeftList").SelectTxt(ReportTagName[i]).Exe();
                }
                catch (Exception ex)
                {
                    i--;
                }
            }
            api.ById("torightBtn").Click();
            api.ById("tagsRightList").SelectTxt(ReportTagName[0]).Exe();

            api.ByXpath("(//input[@name='attachFormat'])[2]").Click();
            api.ByName("emailto").Clear();
            api.ByName("emailto").Enter(sUserEmail).Exe();
            api.ByName("cfgSubmit").Click();
        }

        private void Create_SelfDefined_ODBCData_ExcelReport(string sUserEmail)
        {
            //api.SwitchToCurWindow(0);
            //api.SwitchToFrame("rightFrame", 0);
            //api.ByXpath("//a[contains(@href, '/broadWeb/WaExlViewer/WaExlViewer.asp')]").Click();
            api.ByXpath("//a[contains(@href, 'addRptCfg.aspx')]").Click();

            api.ByName("rptName").Clear();
            api.ByName("rptName").Enter("ExcelReport_SelfDefined_ODBCData_Excel").Exe();
            api.ByXpath("(//input[@name='dataSrc'])[2]").Click();   // ODBC
            api.ByName("selectTemplate").SelectVal("template1.xlsx").Exe();
            switch (slanguage)
            {
                case "ENG":
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
                case "CHT":
                    api.ByName("speTimeFmt").SelectTxt("自訂").Exe();
                    break;
                case "CHS":
                    api.ByName("speTimeFmt").SelectTxt("自订").Exe();
                    break;
                case "JPN":
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
                case "KRN":
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
                case "FRN":

                default:
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
            }

            // set today as start/end date
            string sToday = string.Format("{0:dd}", DateTime.Now);
            string sTomorrow = string.Format("{0:dd}", DateTime.Now.AddHours(8).AddDays(+1));
            int iToday = Int32.Parse(sToday);   // 為了讓讀出來的日期去掉第一個零 ex: "06" -> "6"
            int iTomorrow = Int32.Parse(sTomorrow);
            string ssToday = string.Format("{0}", iToday);
            string ssTomorrow = string.Format("{0}", iTomorrow);
            api.ByName("tStart").Click();
            Thread.Sleep(500);
            api.ByTxt(ssToday).Click();
            Thread.Sleep(500);
            api.ByXpath("(//button[@type='button'])[2]").Click();

            if (iTomorrow != 1)
            {
                api.ByName("tEnd").Click();
                Thread.Sleep(500);
                api.ByTxt(ssTomorrow).Click();
                Thread.Sleep(500);
                //api.ByXpath("//button[@type='button']").Click();    //Click "Now" button
                api.ByXpath("(//button[@type='button'])[2]").Click();
            }
            else
            {
                api.ByName("tEnd").Click();
                Thread.Sleep(500);
                api.ByCss("span.ui-icon.ui-icon-circle-triangle-e").Click();
                Thread.Sleep(500);
                api.ByTxt(ssTomorrow).Click();
                Thread.Sleep(500);
                api.ByXpath("(//button[@type='button'])[2]").Click();
            }

            api.ByName("interval").Clear();
            api.ByName("interval").Enter("1").Exe();
            api.ByXpath("(//input[@name='fileType'])[2]").Click();

            // 湊到最大上限32個TAG
            string[] ReportTagName = { "Calc_ConAna", "Calc_ConDis", "Calc_ModBusAI", "Calc_ModBusAO", "Calc_ModBusDI", "Calc_ModBusDO", "Calc_OPCDA", "Calc_OPCUA","Calc_System",
                                     "Calc_ANDConDisAll", "Calc_ANDDIAll", "Calc_ANDDOAll", "Calc_AvgAIAll", "Calc_AvgAOAll", "Calc_AvgConAnaAll", "Calc_AvgSystemAll",
                                     "AT_AI0050", "AT_AI0100", "AT_AI0150", "AT_AO0050", "AT_AO0100", "AT_AO0150", "AT_DI0050", "AT_DI0100", "AT_DI0150", "AT_DO0050",
                                     "AT_DO0100", "AT_DO0150", "ConAna_0100", "ConDis_0100", "SystemSec_0100", "SystemSec_0200"};
            for (int i = 0; i < ReportTagName.Length; i++)
            {
                try
                {
                    api.ById("tagsLeftList").SelectTxt(ReportTagName[i]).Exe();
                }
                catch (Exception ex)
                {
                    i--;
                }
            }
            api.ById("torightBtn").Click();
            api.ById("tagsRightList").SelectTxt(ReportTagName[0]).Exe();

            api.ByXpath("(//input[@name='attachFormat'])[2]").Click();
            api.ByName("emailto").Clear();
            api.ByName("emailto").Enter(sUserEmail).Exe();
            api.ByName("cfgSubmit").Click();
        }

        private void Create_SelfDefined_Alarm_ExcelReport(string sUserEmail)
        {
            api.ByXpath("//a[contains(@href, 'addRptCfg.aspx')]").Click();

            api.ByName("rptName").Clear();
            api.ByName("rptName").Enter("ExcelReport_SelfDefined_Alarm_Excel").Exe();
            api.ByXpath("(//input[@name='dataSrc'])[3]").Click();   // Alarm
            api.ByName("selectTemplate").SelectVal("AlarmTemplate1_ver.xlsx").Exe();
            switch (slanguage)
            {
                case "ENG":
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
                case "CHT":
                    api.ByName("speTimeFmt").SelectTxt("自訂").Exe();
                    break;
                case "CHS":
                    api.ByName("speTimeFmt").SelectTxt("自订").Exe();
                    break;
                case "JPN":
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
                case "KRN":
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
                case "FRN":

                default:
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
            }

            // set today as start/end date
            string sToday = string.Format("{0:dd}", DateTime.Now);
            string sTomorrow = string.Format("{0:dd}", DateTime.Now.AddHours(8).AddDays(+1));
            int iToday = Int32.Parse(sToday);   // 為了讓讀出來的日期去掉第一個零 ex: "06" -> "6"
            int iTomorrow = Int32.Parse(sTomorrow);
            string ssToday = string.Format("{0}", iToday);
            string ssTomorrow = string.Format("{0}", iTomorrow);
            api.ByName("tStart").Click();
            Thread.Sleep(500);
            api.ByTxt(ssToday).Click();
            Thread.Sleep(500);
            api.ByXpath("(//button[@type='button'])[2]").Click();

            if (iTomorrow != 1)
            {
                api.ByName("tEnd").Click();
                Thread.Sleep(500);
                api.ByTxt(ssTomorrow).Click();
                Thread.Sleep(500);
                //api.ByXpath("//button[@type='button']").Click();    //Click "Now" button
                api.ByXpath("(//button[@type='button'])[2]").Click();
            }
            else
            {
                api.ByName("tEnd").Click();
                Thread.Sleep(500);
                api.ByCss("span.ui-icon.ui-icon-circle-triangle-e").Click();
                Thread.Sleep(500);
                api.ByTxt(ssTomorrow).Click();
                Thread.Sleep(500);
                api.ByXpath("(//button[@type='button'])[2]").Click();
            }

            // 湊到最大上限32個TAG
            string[] ReportTagName = { "AT_AI0001", "AT_AO0001", "AT_DI0001", "AT_DO0001", "Calc_ConAna", "Calc_ConDis", "ConDis_0001", "SystemSec_0001",
                                       "ConAna_0001", "ConAna_0010", "ConAna_0020", "ConAna_0030", "ConAna_0040", "ConAna_0050", "ConAna_0060", "ConAna_0070",
                                       "ConAna_0080", "ConAna_0090", "ConAna_0100", "ConAna_0110", "ConAna_0120", "ConAna_0130", "ConAna_0140", "ConAna_0150", 
                                       "ConAna_0160", "ConAna_0170","ConAna_0180", "ConAna_0190", "ConAna_0200", "ConAna_0210", "ConAna_0220", "ConAna_0230"};
            for (int i = 0; i < ReportTagName.Length; i++)
            {
                try
                {
                    api.ById("tagsLeftList").SelectTxt(ReportTagName[i]).Exe();
                }
                catch (Exception ex)
                {
                    i--;
                }
            }
            api.ById("torightBtn").Click();
            api.ById("tagsRightList").SelectTxt(ReportTagName[0]).Exe();

            api.ByXpath("(//input[@name='attachFormat'])[2]").Click();
            api.ByName("emailto").Clear();
            api.ByName("emailto").Enter(sUserEmail).Exe();
            api.ByName("cfgSubmit").Click();
        }

        private void Create_SelfDefined_ActionLog_ExcelReport(string sUserEmail)
        {
            api.ByXpath("//a[contains(@href, 'addRptCfg.aspx')]").Click();

            api.ByName("rptName").Clear();
            api.ByName("rptName").Enter("ExcelReport_SelfDefined_ActionLog_Excel").Exe();
            api.ByXpath("(//input[@name='dataSrc'])[4]").Click();   // Action Log
            api.ByName("selectTemplate").SelectVal("ActionTemplate1_ver.xlsx").Exe();
            switch (slanguage)
            {
                case "ENG":
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
                case "CHT":
                    api.ByName("speTimeFmt").SelectTxt("自訂").Exe();
                    break;
                case "CHS":
                    api.ByName("speTimeFmt").SelectTxt("自订").Exe();
                    break;
                case "JPN":
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
                case "KRN":
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
                case "FRN":

                default:
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
            }

            // set today as start/end date
            string sToday = string.Format("{0:dd}", DateTime.Now);
            string sTomorrow = string.Format("{0:dd}", DateTime.Now.AddHours(8).AddDays(+1));
            int iToday = Int32.Parse(sToday);   // 為了讓讀出來的日期去掉第一個零 ex: "06" -> "6"
            int iTomorrow = Int32.Parse(sTomorrow);
            string ssToday = string.Format("{0}", iToday);
            string ssTomorrow = string.Format("{0}", iTomorrow);
            api.ByName("tStart").Click();
            Thread.Sleep(500);
            api.ByTxt(ssToday).Click();
            Thread.Sleep(500);
            api.ByXpath("(//button[@type='button'])[2]").Click();

            if (iTomorrow != 1)
            {
                api.ByName("tEnd").Click();
                Thread.Sleep(500);
                api.ByTxt(ssTomorrow).Click();
                Thread.Sleep(500);
                //api.ByXpath("//button[@type='button']").Click();    //Click "Now" button
                api.ByXpath("(//button[@type='button'])[2]").Click();
            }
            else
            {
                api.ByName("tEnd").Click();
                Thread.Sleep(500);
                api.ByCss("span.ui-icon.ui-icon-circle-triangle-e").Click();
                Thread.Sleep(500);
                api.ByTxt(ssTomorrow).Click();
                Thread.Sleep(500);
                api.ByXpath("(//button[@type='button'])[2]").Click();
            }

            api.ByXpath("(//input[@name='attachFormat'])[2]").Click();
            api.ByName("emailto").Clear();
            api.ByName("emailto").Enter(sUserEmail).Exe();
            api.ByName("cfgSubmit").Click();
        }

        private void Create_SelfDefined_EventLog_ExcelReport(string sUserEmail)
        {
            api.ByXpath("//a[contains(@href, 'addRptCfg.aspx')]").Click();

            api.ByName("rptName").Clear();
            api.ByName("rptName").Enter("ExcelReport_SelfDefined_EventLog_Excel").Exe();
            api.ByXpath("(//input[@name='dataSrc'])[5]").Click();   // Action Log
            api.ByName("selectTemplate").SelectVal("EventTemplate1_ver.xlsx").Exe();
            switch (slanguage)
            {
                case "ENG":
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
                case "CHT":
                    api.ByName("speTimeFmt").SelectTxt("自訂").Exe();
                    break;
                case "CHS":
                    api.ByName("speTimeFmt").SelectTxt("自订").Exe();
                    break;
                case "JPN":
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
                case "KRN":
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
                case "FRN":

                default:
                    api.ByName("speTimeFmt").SelectTxt("Self-Defined").Exe();
                    break;
            }

            // set today as start/end date
            string sToday = string.Format("{0:dd}", DateTime.Now);
            string sTomorrow = string.Format("{0:dd}", DateTime.Now.AddHours(8).AddDays(+1));
            int iToday = Int32.Parse(sToday);   // 為了讓讀出來的日期去掉第一個零 ex: "06" -> "6"
            int iTomorrow = Int32.Parse(sTomorrow);
            string ssToday = string.Format("{0}", iToday);
            string ssTomorrow = string.Format("{0}", iTomorrow);
            api.ByName("tStart").Click();
            Thread.Sleep(500);
            api.ByTxt(ssToday).Click();
            Thread.Sleep(500);
            api.ByXpath("(//button[@type='button'])[2]").Click();

            if (iTomorrow != 1)
            {
                api.ByName("tEnd").Click();
                Thread.Sleep(500);
                api.ByTxt(ssTomorrow).Click();
                Thread.Sleep(500);
                //api.ByXpath("//button[@type='button']").Click();    //Click "Now" button
                api.ByXpath("(//button[@type='button'])[2]").Click();
            }
            else
            {
                api.ByName("tEnd").Click();
                Thread.Sleep(500);
                api.ByCss("span.ui-icon.ui-icon-circle-triangle-e").Click();
                Thread.Sleep(500);
                api.ByTxt(ssTomorrow).Click();
                Thread.Sleep(500);
                api.ByXpath("(//button[@type='button'])[2]").Click();
            }

            api.ByXpath("(//input[@name='attachFormat'])[2]").Click();
            api.ByName("emailto").Clear();
            api.ByName("emailto").Enter(sUserEmail).Exe();
            api.ByName("cfgSubmit").Click();
        }

        private void ReturnSCADAPage()
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMainRight.asp') and contains(@href, 'name=TestSCADA')]").Click();

        }

        private void Start_Click(object sender, EventArgs e)
        {
            long lErrorCode = (long)ErrorCode.SUCCESS;
            EventLog.AddLog("===Create Excel Report start===");
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            EventLog.AddLog("User Email Address= " + UserEmail.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text, UserEmail.Text);
            EventLog.AddLog("===Create Excel Report end===");
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
            /*
            tpc.F_WritePrivateProfileString("ProjectName", "Ground PC or Primary PC", "TestProject", @"C:\WebAccessAutoTestSetting.ini");
            tpc.F_WritePrivateProfileString("ProjectName", "Cloud PC or Backup PC", "CTestProject", @"C:\WebAccessAutoTestSetting.ini");
            tpc.F_WritePrivateProfileString("IP", "Ground PC or Primary PC", "172.18.3.62", @"C:\WebAccessAutoTestSetting.ini");
            tpc.F_WritePrivateProfileString("IP", "Cloud PC or Backup PC", "172.18.3.65", @"C:\WebAccessAutoTestSetting.ini");
            */
            tpc.F_GetPrivateProfileString("UserInfo", "Language", "NA", sDefaultUserLanguage, 255, sFilePath);
            tpc.F_GetPrivateProfileString("UserInfo", "Email", "NA", sDefaultUserEmail, 255, sFilePath);
            tpc.F_GetPrivateProfileString("ProjectName", "Ground PC or Primary PC", "NA", sDefaultProjectName1, 255, sFilePath);
            tpc.F_GetPrivateProfileString("ProjectName", "Cloud PC or Backup PC", "NA", sDefaultProjectName2, 255, sFilePath);
            tpc.F_GetPrivateProfileString("IP", "Ground PC or Primary PC", "NA", sDefaultIP1, 255, sFilePath);
            tpc.F_GetPrivateProfileString("IP", "Cloud PC or Backup PC", "NA", sDefaultIP2, 255, sFilePath);
            slanguage = sDefaultUserLanguage.ToString();    // 在這邊讀取使用語言
            UserEmail.Text = sDefaultUserEmail.ToString();
            ProjectName.Text = sDefaultProjectName1.ToString();
            WebAccessIP.Text = sDefaultIP1.ToString();
        }

        private void CheckifIniFileChange()
        {
            StringBuilder sDefaultUserEmail = new StringBuilder(255);
            StringBuilder sDefaultProjectName1 = new StringBuilder(255);
            StringBuilder sDefaultProjectName2 = new StringBuilder(255);
            StringBuilder sDefaultIP1 = new StringBuilder(255);
            StringBuilder sDefaultIP2 = new StringBuilder(255);
            if (System.IO.File.Exists(sIniFilePath))    // 比對ini檔與ui上的值是否相同
            {
                EventLog.AddLog(".ini file exist, check if .ini file need to update");
                tpc.F_GetPrivateProfileString("UserInfo", "Email", "NA", sDefaultUserEmail, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("ProjectName", "Ground PC or Primary PC", "NA", sDefaultProjectName1, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("ProjectName", "Cloud PC or Backup PC", "NA", sDefaultProjectName2, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("IP", "Ground PC or Primary PC", "NA", sDefaultIP1, 255, sIniFilePath);
                tpc.F_GetPrivateProfileString("IP", "Cloud PC or Backup PC", "NA", sDefaultIP2, 255, sIniFilePath);

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
            }
            else
            {
                EventLog.AddLog(".ini file not exist, create new .ini file. Path: " + sIniFilePath);
                tpc.F_WritePrivateProfileString("UserInfo", "Email", UserEmail.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("ProjectName", "Ground PC or Primary PC", ProjectName.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("ProjectName", "Cloud PC or Backup PC", "CTestProject", sIniFilePath);
                tpc.F_WritePrivateProfileString("IP", "Ground PC or Primary PC", WebAccessIP.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("IP", "Cloud PC or Backup PC", "172.18.3.65", sIniFilePath);
            }
        }

    }
}