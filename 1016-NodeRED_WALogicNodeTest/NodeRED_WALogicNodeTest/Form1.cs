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
using System.IO;

namespace NodeRED_WALogicNodeTest
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
            EventLog.AddLog("===NodeRED_WALogicNode test start (by iATester)===");

            if (System.IO.File.Exists(sIniFilePath))    // 再load一次
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===NodeRED_WALogicNode test end (by iATester)===");

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

            // Launch Firefox and login
            api.LinkWebUI(baseUrl + "/broadWeb/bwRoot.asp?username=admin");
            api.ByXpath("//a[contains(@href, '/broadWeb/bwconfig.asp?username=admin')]").Click();
            api.ById("userField").Enter("").Submit().Exe();
            PrintStep("Login WebAccess");

            // Configure project by project name
            api.ByXpath("(//a[contains(@href, 'userName=admin&projectName1=" + sProjectName + "')])[2]").Click();   //跳到NodeRED設定頁面
            PrintStep("Configure NodeRED page");

            // Logic NodeRED test
            /*
            主要動作
             * 1 新增第2個分頁
             * 2 在第2個分頁導入node red資料
             * 3 新增第3個分頁 刪除第3個分頁 (因為導入資料後 無法作滑鼠點擊的動作 只好用這個動作取代滑鼠點擊)
             * 4 第2個分頁重新命名 並 deploy
             * 5 觸發且讀取得到的值並判斷是否正確
             * 6 刪除第2個分頁
             */
            bool bNodeREDResult = bNodeREDLogicNodeTest(sProjectName);

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

            if (bSeleniumResult && bNodeREDResult)
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

        private bool bNodeREDLogicNodeTest(string sProjectName)
        {
            bool bTestResult = true;
            string[] sLogicNodeName = { "And", "Compare", "Equal", "GreaterOrEqual", "GreaterThan", "LessOrEqual", "LessThan", "NAND-Gate", "NOR-Gate", "Not", "NotEqual", "Or", "XOR-Gate" };
            string[] sOutputType = { "boolean", "number", "boolean", "boolean", "boolean", "boolean", "boolean", "boolean", "boolean", "boolean", "boolean", "boolean", "boolean" };
            string[] sOutputValue = { "false", "1", "false", "true", "true", "false", "false", "false", "false", "false", "true", "true", "false" };

            for (int i = 1; i <= sLogicNodeName.Length; i++)
            {
                //Step1: add workspace
                api.ByXpath("//a[contains(@href, '#debug')]").Click();      //切到debug視窗
                api.ByXpath("//a[@id='btn-workspace-add-tab']/i").Click();  //新增Sheet2 (主要操作區)

                //Step2: import test case
                string sCurrentFilePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetAssembly(this.GetType()).Location);
                string sourceFile = sCurrentFilePath + string.Format("\\NodeREDSample\\{0}.txt", sLogicNodeName[i - 1]);
                StreamReader sr = new StreamReader(sourceFile, Encoding.Default);
                string line;
                string sNodeRED_Sample = "";
                while ((line = sr.ReadLine()) != null)
                {
                    sNodeRED_Sample += line.ToString();
                    //Console.WriteLine(line.ToString());
                }
                //Replace project name here
                sNodeRED_Sample = sNodeRED_Sample.Replace("CTestProject", sProjectName);

                /* this method doesn't work...WTF
                api.ByXpath("//a[@id='btn-sidemenu']/i").Click(); // click side menu
                api.ByXpath("//a[@id='menu-item-import']").MoveToEle(); // moveToElement
                Thread.Sleep(2000);
                api.ById("menu-item-import-clipboard").Click(); // click Import
                Thread.Sleep(2000);
                 * */
                SendKeys.SendWait("^{i}");  // import tag
                Thread.Sleep(2000);
                api.ById("clipboard-import").Enter(sNodeRED_Sample).Exe();
                api.ById("clipboard-dialog-ok").Click();    // click ok

                api.ByXpath("//a[@id='btn-workspace-add-tab']/i").Click();  //新增Sheet3
                Thread.Sleep(500);
                api.ByXpath("//div[2]/ul/li[3]/a").DoubleClick();   //點擊Sheet3 -> 為了使貼上去的資料固定住
                api.ByXpath("//button[@type='button']").Click();    //刪除Sheet3

                api.ByXpath("//div[2]/ul/li[2]/a").Click();            //點擊Sheet2
                Thread.Sleep(2000);
                api.ByXpath("//div[2]/ul/li[2]/a").DoubleClick();
                Thread.Sleep(2000);
                api.ById("node-input-workspace-name").Clear();
                Thread.Sleep(2000);
                api.ById("node-input-workspace-name").Enter(sLogicNodeName[i - 1]).Exe();
                Thread.Sleep(2000);
                api.ByXpath("(//button[@type='button'])[2]").Click();
                Thread.Sleep(2000);

                //Step3: Deploy setting
                api.ByXpath("//a[@id='btn-deploy']/span").Click();  // deploy
                Thread.Sleep(3000);

                //Step4: Trigger
                api.ByCss("g.node_button.node_left_button > rect.node_button_button").Click();

                //Step5: catch debug message and judgment
                string sDebugMessageDate = api.ByCss("#debug-content > div:nth-child(1) > span.debug-message-date").GetText();
                string sDebugMessageName = api.ByCss("#debug-content > div > span.debug-message-name").GetText();
                string sDebugMessageTopic = api.ByCss("#debug-content > div > span.debug-message-topic").GetText();
                string sDebugMessagePayload = api.ByCss("#debug-content > div > span.debug-message-payload").GetText();

                EventLog.AddLog(sLogicNodeName[i - 1] + " conversioin test");
                EventLog.AddLog("Debug Message:");
                EventLog.AddLog(sDebugMessageDate + " " + sDebugMessageName);
                EventLog.AddLog(sDebugMessageTopic);
                EventLog.AddLog(sDebugMessagePayload);
                string[] sType = sDebugMessageTopic.Split(new string[] { ": " }, StringSplitOptions.RemoveEmptyEntries); // 切割文字抓取回傳type

                if (sType[1] != sOutputType[i - 1])
                {
                    bTestResult = false;
                    EventLog.AddLog("Output type ERROR!!");
                    EventLog.AddLog("Correct type is: " + sOutputType[i - 1]);
                }

                if (sDebugMessagePayload != sOutputValue[i - 1])
                {
                    bTestResult = false;
                    EventLog.AddLog("Output value ERROR!!");
                    EventLog.AddLog("Correct value is: " + sOutputValue[i - 1]);
                }

                //Step6: clear debug message
                api.ByXpath("//a[@id='debug-tab-clear']/i").Click();    // clear debug message

                //Step7: delete test sheet
                api.ByXpath("//div[2]/ul/li[2]/a").DoubleClick();       //點擊Sheet2
                api.ByXpath("//button[@type='button']").Click();        //刪除Sheet2
                api.ByXpath("(//button[@type='button'])[4]").Click();   //確認刪除Sheet2
                api.ByXpath("//a[@id='btn-deploy']/span").Click();      // deploy

                PrintStep(sLogicNodeName[i - 1] + " Test");
            }

            return bTestResult;
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

        private void SetupBasicAnalogTagConfig()
        {
            api.ByName("Datalog").Click();
            api.ByName("DataLogDB").Clear();
            api.ByName("DataLogDB").Enter("0").Exe();
            api.ByName("ChangeLog").Click();
            api.ByName("SpanHigh").Clear();
            api.ByName("SpanHigh").Enter("1000").Exe();
            api.ByName("OutputHigh").Clear();
            api.ByName("OutputHigh").Enter("1000").Exe();
            api.ByName("ReservedInt1").SelectTxt("2").Exe();
            api.ByXpath("(//input[@name='LogTmRadio'])[2]").Click();
        }

        private void SetupBasicDigitalTagConfig()
        {
            api.ByName("Datalog").Click();
            api.ByName("DataLogDB").Clear();
            api.ByName("DataLogDB").Enter("0").Exe();
            api.ByName("ChangeLog").Click();
            api.ByName("ReservedInt1").Click();
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
            EventLog.AddLog("===NodeRED_WALogicNode test start===");
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===NodeRED_WALogicNode test end===");
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

    }
}
