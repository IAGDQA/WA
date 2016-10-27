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
using System.IO;
using ThirdPartyToolControl;
using iATester;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInOut
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
            EventLog.AddLog("===Excel in or out start (by iATester)===");
            if (System.IO.File.Exists(sIniFilePath))    // 再load一次
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===Excel in or out end (by iATester)===");

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
            api.LinkWebUI(baseUrl + "/broadWeb/bwconfig.asp?username=admin");

            api.ById("userField").Enter("").Submit().Exe();
            PrintStep("Login WebAccess");

            // Configure project by project name
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
            PrintStep("Configure project");


            //Case1: Excel in
            EventLog.AddLog("Excel in...");
            //string sSourceFile = "C:\\WALogData\\bwTagImport_AutoTest"; //debug
            string sCurrentFilePath = Directory.GetCurrentDirectory();
            string sSourceFile = sCurrentFilePath + "\\ExcelIn\\bwTagImport_AutoTest";

            EventLog.AddLog("Set project name to excel file");
            SetExcelProjectName(sProjectName, sSourceFile);

            ExcuteExcelIn(sSourceFile);
            Thread.Sleep(5000);
            string fileNameTar_in = string.Format("ExcelIn_{0:yyyyMMdd_hhmmss}", DateTime.Now);
            PrintScreen(fileNameTar_in, sTestLogFolder);
            PrintStep("Excel in");

            api.Refresh();
            ReturnSCADAPage();

            //Case2: Excel out
            EventLog.AddLog("Excel out...");
            string sdestFile = sTestLogFolder + string.Format("\\bwTagExport_{0:yyyyMMdd_hhmmss}", DateTime.Now);
            ExcuteExcelOut(sdestFile);
            Thread.Sleep(5000);
            string fileNameTar_out = string.Format("ExcelOut_{0:yyyyMMdd_hhmmss}", DateTime.Now);
            PrintScreen(fileNameTar_out, sTestLogFolder);
            PrintStep("Excel out");

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

        private void SetExcelProjectName(string sProjectName, string sSourceFile)
        {
            //設定必要的物件
            //按照順序分別是Application > Workbook > Worksheet > Range > Cell
            //(1) Application ：代表一個 Excel 程序。
            //(2) WorkBook ：代表一個 Excel 工作簿。
            //(3) WorkSheet ：代表一個 Excel 工作表，一個 WorkBook 包含好幾個工作表。
            //(4) Range ：代表 WorkSheet 中的多個單元格區域。
            //(5) Cell ：代表 WorkSheet 中的一個單元格。
            Excel.Application App = new Excel.Application();

            //取得欲寫入的檔案路徑
            string strPath = sSourceFile + ".XLS";
            Excel.Workbook Wbook = App.Workbooks.Open(strPath);

            //將欲修改的檔案屬性設為非唯讀(Normal)，若寫入檔案為唯讀，則會無法寫入
            System.IO.FileInfo xlsAttribute = new FileInfo(strPath);
            xlsAttribute.Attributes = FileAttributes.Normal;


            Excel.Worksheet Wsheet = (Excel.Worksheet)Wbook.Sheets["BwAnalog"];
            //取得工作表的單元格
            for (int i = 2; i <= 1501; i++)
            {
                Excel.Range aRangeChange = Wsheet.get_Range("A" + i.ToString());

                //在工作表的特定儲存格，設定內容
                aRangeChange.Value2 = sProjectName;
            }

            Excel.Worksheet Wsheet2 = (Excel.Worksheet)Wbook.Sheets["BwDiscrete"];
            //取得工作表的單元格
            for (int i = 2; i <= 751; i++)
            {
                Excel.Range aRangeChange = Wsheet2.get_Range("A" + i.ToString());

                //在工作表的特定儲存格，設定內容
                aRangeChange.Value2 = sProjectName;
            }

            Excel.Worksheet Wsheet3 = (Excel.Worksheet)Wbook.Sheets["BwText"];
            //取得工作表的單元格
            for (int i = 2; i <= 251; i++)
            {
                Excel.Range aRangeChange = Wsheet3.get_Range("A" + i.ToString());

                //在工作表的特定儲存格，設定內容
                aRangeChange.Value2 = sProjectName;
            }

            Excel.Worksheet Wsheet4 = (Excel.Worksheet)Wbook.Sheets["BwCalcAnalog"];
            //取得工作表的單元格
            for (int i = 2; i <= 92; i++)
            {
                Excel.Range aRangeChange = Wsheet4.get_Range("A" + i.ToString());

                //在工作表的特定儲存格，設定內容
                aRangeChange.Value2 = sProjectName;
            }

            Excel.Worksheet Wsheet5 = (Excel.Worksheet)Wbook.Sheets["BwCalcDiscrete"];
            //取得工作表的單元格
            for (int i = 2; i <= 46; i++)
            {
                Excel.Range aRangeChange = Wsheet5.get_Range("A" + i.ToString());

                //在工作表的特定儲存格，設定內容
                aRangeChange.Value2 = sProjectName;
            }

            Excel.Worksheet Wsheet6 = (Excel.Worksheet)Wbook.Sheets["BwAcc"];
            //取得工作表的單元格
            for (int i = 2; i <= 251; i++)
            {
                Excel.Range aRangeChange = Wsheet6.get_Range("A" + i.ToString());

                //在工作表的特定儲存格，設定內容
                aRangeChange.Value2 = sProjectName;
            }

            Excel.Worksheet Wsheet7 = (Excel.Worksheet)Wbook.Sheets["BwAlarmAnalog"];
            //取得工作表的單元格
            for (int i = 2; i <= 255; i++)
            {
                Excel.Range aRangeChange = Wsheet7.get_Range("A" + i.ToString());

                //在工作表的特定儲存格，設定內容
                aRangeChange.Value2 = sProjectName;
            }

            Excel.Worksheet Wsheet8 = (Excel.Worksheet)Wbook.Sheets["BwAlarmDiscrete"];
            //取得工作表的單元格
            for (int i = 2; i <= 5; i++)
            {
                Excel.Range aRangeChange = Wsheet8.get_Range("A" + i.ToString());

                //在工作表的特定儲存格，設定內容
                aRangeChange.Value2 = sProjectName;
            }

            //設置禁止彈出保存和覆蓋的詢問提示框
            Wsheet.Application.DisplayAlerts = false;
            Wsheet.Application.AlertBeforeOverwriting = false;

            //保存工作表，因為禁止彈出儲存提示框，所以需在此儲存，否則寫入的資料會無法儲存
            Wbook.Save();

            //關閉EXCEL
            Wbook.Close();

            //離開應用程式
            App.Quit();
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

        private void ExcuteExcelIn(string sSourceFile)
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/odbc/odbcPg1.asp?pos=import')]").Click();
            api.ByName("XlsName").Clear();
            api.ByName("XlsName").Enter(sSourceFile).Submit().Exe();
        }

        private void ExcuteExcelOut(string sdestFile)
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/odbc/odbcPg1.asp?pos=export')]").Click();
            api.ByName("XlsName").Clear();
            api.ByName("XlsName").Enter(sdestFile).Submit().Exe();
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

        private void ReturnSCADAPage()
        {
            //driver.SwitchTo().Window(driver.CurrentWindowHandle);   // Return parent frame
            //driver.SwitchTo().Frame("leftFrame");                   // Focus on left frame
            //driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwMainRight.asp') and contains(@href, 'name=TestSCADA')]")).Click();
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMainRight.asp') and contains(@href, 'name=TestSCADA')]").Click();

        }

        private void Start_Click(object sender, EventArgs e)
        {
            long lErrorCode = (long)ErrorCode.SUCCESS;
            EventLog.AddLog("===Excel in or out start===");
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===Excel in or out end===");
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

        private void ProjectName_TextChanged(object sender, EventArgs e)
        {

        }

        private void WebAccessIP_TextChanged(object sender, EventArgs e)
        {

        }

        private void TestLogFolder_TextChanged(object sender, EventArgs e)
        {

        }

        private void Browser_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

    }
}

