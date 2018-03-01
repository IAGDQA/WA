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
using System.IO;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelInOut
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
        string sTestItemName = "ExcelInOut";
        string sIniFilePath = @"C:\WebAccessAutoTestSettingInfo.ini";
        string sTestLogFolder = @"C:\WALogData";
        int[] getTheNumberofImportTag = new int[8];

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
                    Thread.Sleep(2000);
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

            //Case1: Excel in
            if (bPartResult == true)
            {
                EventLog.AddLog("Excel in");
                sw.Reset(); sw.Start();
                try
                {
                    //string sSourceFile = @"C:\WALogData\bwTagImport_AutoTest"; //debug
                    //string sCurrentFilePath = Directory.GetCurrentDirectory();
                    string sCurrentFilePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetAssembly(this.GetType()).Location);
                    string sSourceFile = sCurrentFilePath + "\\ExcelIn\\bwTagImport_AutoTest";

                    EventLog.AddLog("Set project name to excel file");
                    SetExcelProjectName(sPrimaryProject, sSourceFile);
                    EventLog.AddLog("Set project name to excel file done!");

                    ExcuteExcelIn(sSourceFile);
                    //Thread.Sleep(20000);
                    string fileNameTar_in = string.Format("ExcelIn_{0:yyyyMMdd_hhmmss}", DateTime.Now);
                    EventLog.PrintScreen(fileNameTar_in);
                }
                catch (Exception ex)
                {
                    EventLog.AddLog(@"Error occurred Excel in: " + ex.ToString());
                    bPartResult = false;
                }
                sw.Stop();
                PrintStep("Excel in", "Excel in", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());

                Thread.Sleep(1000);
            }

            //Case2: Excel out
            if (bPartResult == true)
            {
                EventLog.AddLog("Excel out");
                sw.Reset(); sw.Start();
                try
                {
                    string sdestFile = sTestLogFolder + string.Format("\\bwTagExport_{0:yyyyMMdd_hhmmss}", DateTime.Now);
                    ExcuteExcelOut(sdestFile);
                    //Thread.Sleep(20000);
                    string fileNameTar_out = string.Format("ExcelOut_{0:yyyyMMdd_hhmmss}", DateTime.Now);
                    EventLog.PrintScreen(fileNameTar_out);
                }
                catch (Exception ex)
                {
                    EventLog.AddLog(@"Error occurred Excel out: " + ex.ToString());
                    bPartResult = false;
                }
                sw.Stop();
                PrintStep("Excel out", "Excel out", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());

                Thread.Sleep(1000);
            }
            Thread.Sleep(1000);

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
            EventLog.AddLog("Get original excel file path: " + strPath);

            //將欲修改的檔案屬性設為非唯讀(Normal)，若寫入檔案為唯讀，則會無法寫入
            EventLog.AddLog("Set the file attribute to readable type");
            System.IO.FileInfo xlsAttribute = new FileInfo(strPath);
            xlsAttribute.Attributes = FileAttributes.Normal;

            EventLog.AddLog("Set BwAnalog..");
            Excel.Worksheet Wsheet = (Excel.Worksheet)Wbook.Sheets["BwAnalog"];
            //取得工作表的單元格
            for (int i = 2; i <= 9999; i++) // 1502
            {
                Excel.Range aRangeChange = Wsheet.get_Range("A" + i.ToString());
                
                if (aRangeChange.Value2 == "**********")
                {
                    getTheNumberofImportTag[0] = i - 2;
                    EventLog.AddLog(string.Format("The number of import tags is {0}", i-2)); // 去頭去尾
                    break;
                }
                else
                {
                    aRangeChange.Value2 = sProjectName; //在工作表的特定儲存格，設定內容
                }
            }

            EventLog.AddLog("Set BwDiscrete..");
            Excel.Worksheet Wsheet2 = (Excel.Worksheet)Wbook.Sheets["BwDiscrete"];
            //取得工作表的單元格
            for (int i = 2; i <= 9999; i++)  //752
            {
                Excel.Range aRangeChange = Wsheet2.get_Range("A" + i.ToString());

                if (aRangeChange.Value2 == "**********")
                {
                    getTheNumberofImportTag[1] = i - 2;
                    EventLog.AddLog(string.Format("The number of import tags is {0}", i - 2)); // 去頭去尾
                    break;
                }
                else
                {
                    aRangeChange.Value2 = sProjectName; //在工作表的特定儲存格，設定內容
                }
            }

            EventLog.AddLog("Set BwText..");
            Excel.Worksheet Wsheet3 = (Excel.Worksheet)Wbook.Sheets["BwText"];
            //取得工作表的單元格
            for (int i = 2; i <= 9999; i++)  //251
            {
                Excel.Range aRangeChange = Wsheet3.get_Range("A" + i.ToString());

                if (aRangeChange.Value2 == "**********")
                {
                    getTheNumberofImportTag[2] = i - 2;
                    EventLog.AddLog(string.Format("The number of import tags is {0}", i - 2)); // 去頭去尾
                    break;
                }
                else
                {
                    aRangeChange.Value2 = sProjectName; //在工作表的特定儲存格，設定內容
                }
            }

            EventLog.AddLog("Set BwCalcAnalog..");
            Excel.Worksheet Wsheet4 = (Excel.Worksheet)Wbook.Sheets["BwCalcAnalog"];
            //取得工作表的單元格
            for (int i = 2; i <= 9999; i++)   // 92
            {
                Excel.Range aRangeChange = Wsheet4.get_Range("A" + i.ToString());

                if (aRangeChange.Value2 == "**********")
                {
                    getTheNumberofImportTag[3] = i - 2;
                    EventLog.AddLog(string.Format("The number of import tags is {0}", i - 2)); // 去頭去尾
                    break;
                }
                else
                {
                    aRangeChange.Value2 = sProjectName; //在工作表的特定儲存格，設定內容
                }
            }

            EventLog.AddLog("Set BwCalcDiscrete..");
            Excel.Worksheet Wsheet5 = (Excel.Worksheet)Wbook.Sheets["BwCalcDiscrete"];
            //取得工作表的單元格
            for (int i = 2; i <= 9999; i++)   // 46
            {
                Excel.Range aRangeChange = Wsheet5.get_Range("A" + i.ToString());

                if (aRangeChange.Value2 == "**********")
                {
                    getTheNumberofImportTag[4] = i - 2;
                    EventLog.AddLog(string.Format("The number of import tags is {0}", i - 2)); // 去頭去尾
                    break;
                }
                else
                {
                    aRangeChange.Value2 = sProjectName; //在工作表的特定儲存格，設定內容
                }
            }

            EventLog.AddLog("Set BwAcc..");
            Excel.Worksheet Wsheet6 = (Excel.Worksheet)Wbook.Sheets["BwAcc"];
            //取得工作表的單元格
            for (int i = 2; i <= 9999; i++)  // 251
            {
                Excel.Range aRangeChange = Wsheet6.get_Range("A" + i.ToString());

                if (aRangeChange.Value2 == "**********")
                {
                    getTheNumberofImportTag[5] = i - 2;
                    EventLog.AddLog(string.Format("The number of import tags is {0}", i - 2)); // 去頭去尾
                    break;
                }
                else
                {
                    aRangeChange.Value2 = sProjectName; //在工作表的特定儲存格，設定內容
                }
            }

            EventLog.AddLog("Set BwAlarmAnalog..");
            Excel.Worksheet Wsheet7 = (Excel.Worksheet)Wbook.Sheets["BwAlarmAnalog"];
            //取得工作表的單元格
            for (int i = 2; i <= 9999; i++)  //255
            {
                Excel.Range aRangeChange = Wsheet7.get_Range("A" + i.ToString());

                if (aRangeChange.Value2 == "**********")
                {
                    getTheNumberofImportTag[6] = i - 2;
                    EventLog.AddLog(string.Format("The number of import tags is {0}", i - 2)); // 去頭去尾
                    break;
                }
                else
                {
                    aRangeChange.Value2 = sProjectName; //在工作表的特定儲存格，設定內容
                }
            }

            EventLog.AddLog("Set BwAlarmDiscrete..");
            Excel.Worksheet Wsheet8 = (Excel.Worksheet)Wbook.Sheets["BwAlarmDiscrete"];
            //取得工作表的單元格
            for (int i = 2; i <= 9999; i++)    // 5
            {
                Excel.Range aRangeChange = Wsheet8.get_Range("A" + i.ToString());

                if (aRangeChange.Value2 == "**********")
                {
                    getTheNumberofImportTag[7] = i - 2;
                    EventLog.AddLog(string.Format("The number of import tags is {0}", i - 2)); // 去頭去尾
                    break;
                }
                else
                {
                    aRangeChange.Value2 = sProjectName; //在工作表的特定儲存格，設定內容
                }
            }

            //設置禁止彈出保存和覆蓋的詢問提示框
            Wsheet.Application.DisplayAlerts = false;
            Wsheet.Application.AlertBeforeOverwriting = false;

            //保存工作表，因為禁止彈出儲存提示框，所以需在此儲存，否則寫入的資料會無法儲存
            Wbook.Save();
            EventLog.AddLog("Save");

            //關閉EXCEL
            Wbook.Close();
            EventLog.AddLog("Close");

            //離開應用程式
            App.Quit();
            EventLog.AddLog("Quit");
        }

        private void ExcuteExcelIn(string sSourceFile)
        {
            driver.SwitchTo().Frame("rightFrame");
            driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/odbc/odbcPg1.asp?pos=import')]")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Name("XlsName")).Clear();
            driver.FindElement(By.Name("XlsName")).SendKeys(sSourceFile);
            driver.FindElement(By.Name("submit")).Click();
            Thread.Sleep(1500);
            string[] getName = new string[8];
            string[] getDescription = new string[8];
            EventLog.AddLog("The result of WebAccess excel in is below: ");
            for (int i = 0; i < 8; i++)
            {
                getName[i] = driver.FindElement(By.XPath(string.Format("//*[@id='form1']/table/tbody/tr[{0}]/td[3]/font", i + 3))).Text;
                getDescription[i] = driver.FindElement(By.XPath(string.Format("//*[@id='form1']/table/tbody/tr[{0}]/td[4]/font",i+3))).Text;
                EventLog.AddLog( getName[i] + " --> " + getDescription[i] );

                if (!getDescription[i].Contains(getTheNumberofImportTag[i].ToString()))
                {
                    bPartResult = false;
                    EventLog.AddLog("Check the number of import tag is not correct!!");
                    EventLog.AddLog("The number in excel is: " + getTheNumberofImportTag[i].ToString());
                }
            }
            driver.FindElement(By.Name("submit")).Click();
            // Check excel data
            driver.SwitchTo().ParentFrame();
        }

        private void ExcuteExcelOut(string sdestFile)
        {
            driver.SwitchTo().Frame("rightFrame");
            driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/odbc/odbcPg1.asp?pos=export')]")).Click();
            Thread.Sleep(1000);
            driver.FindElement(By.Name("XlsName")).Clear();
            driver.FindElement(By.Name("XlsName")).SendKeys(sdestFile);
            driver.FindElement(By.Name("submit")).Click();
            Thread.Sleep(1500);
            driver.FindElement(By.Name("submit")).Click();

            if (!System.IO.File.Exists(sdestFile+".XLS"))
            {
                bPartResult = false;
                EventLog.AddLog("Error!! Cannot find " + sdestFile + ".XLS !!");
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

