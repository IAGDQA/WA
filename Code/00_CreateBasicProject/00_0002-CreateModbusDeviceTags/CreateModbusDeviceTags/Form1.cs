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
//using System.IO;
//using System.Reflection;
//using Excel = Microsoft.Office.Interop.Excel;

namespace CreateModbusDeviceTags
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
        string sTestItemName = "CreateModbusDeviceTags";
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

            if (bPartResult == true)
            {
                EventLog.AddLog("Create ModBus tag");
                sw.Reset(); sw.Start();
                try
                {
                    //Step0: check comport exist
                    EventLog.AddLog("Check Comport");
                    CheckComport();
                    
                    //Step1: add Comport
                    EventLog.AddLog("Add Comport");
                    AddComport();
                    //PrintStep("Add Comport");

                    Thread.Sleep(1000);

                    //Step2: add Device
                    EventLog.AddLog("Add Device");
                    AddDevice(sPrimaryIP);
                    //PrintStep("Add Device");

                    //Step2: Create Modbus Tag
                    EventLog.AddLog("Create Modbus Tag");
                    CreateModbusTag();
                }
                catch (Exception ex)
                {
                    EventLog.AddLog(@"Error occurred create ModBus tag: " + ex.ToString());
                    bPartResult = false;
                }
                sw.Stop();
                PrintStep("Create", "Create ModBus tag", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());

                Thread.Sleep(1000);
            }

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

        private void CheckComport()
        {
            driver.SwitchTo().Frame("leftFrame");
            if (wcf.IsTestElementPresent(driver, "XPath", "//a[contains(@href, '/broadWeb/bwMainRight.asp?pos=comport') and contains(@href, 'name=3')]"))
            {
                driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwMainRight.asp?pos=comport') and contains(@href, 'name=3')]")).Click();
                Thread.Sleep(1500);
                EventLog.AddLog("Delete the com port");
                driver.SwitchTo().ParentFrame();
                driver.SwitchTo().Frame("rightFrame");
                driver.FindElement(By.XPath("//a[2]/font/b")).Click();
                Thread.Sleep(1000);
                driver.SwitchTo().Alert().Accept();

                long lStartTime = Environment.TickCount;
                long lEndTime = 0;
                do
                {
                    if (!wcf.IsTestElementPresent(driver, "XPath", "//a[2]/font/b"))
                        break;
                    lEndTime = Environment.TickCount;
                } while ((lEndTime - lStartTime) <= 30000);

                if (lEndTime - lStartTime > 30000)
                {
                    bPartResult = false;
                    EventLog.AddLog("Cannot delete com port in 30s, test fail");
                }
                Thread.Sleep(3000);
                driver.Navigate().Refresh();
                Thread.Sleep(5000);
            }

            driver.SwitchTo().ParentFrame();
        }

        private void AddComport()
        {
            driver.SwitchTo().Frame("rightFrame");
            //driver.FindElement(By.XPath("//a[3]/font/b")).Click();
            driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/comport/comportPg.asp?pos=node')]")).Click();
            Thread.Sleep(500);
            new SelectElement(driver.FindElement(By.Name("infName"))).SelectByText("TCPIP");    // maybe select by value
            Thread.Sleep(500);
            driver.FindElement(By.Name("ComportNbr")).Clear();
            driver.FindElement(By.Name("ComportNbr")).SendKeys("3");  // comport = 3
            driver.FindElement(By.Name("submit")).Click();
        }

        private void AddDevice(string sWebAccessIP)
        {
            driver.SwitchTo().ParentFrame();
            driver.SwitchTo().Frame("leftFrame");

            driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwMainRight.asp?pos=comport') and contains(@href, 'name=3')]")).Click(); // port = 3
            driver.SwitchTo().ParentFrame();
            driver.SwitchTo().Frame("rightFrame");
            driver.FindElement(By.XPath("//a[3]/font/b")).Click();
            //driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/device/devPg.asp')] and contains(@href, 'action=add')]")).Click();
            driver.FindElement(By.Name("DeviceName")).Clear();
            driver.FindElement(By.Name("DeviceName")).SendKeys("ModSim");
            new SelectElement(driver.FindElement(By.Name("DeviceTypeSel"))).SelectByText("Modicon");
            driver.FindElement(By.Name("P_IPAddr")).Clear();
            driver.FindElement(By.Name("P_IPAddr")).SendKeys(sWebAccessIP);
            driver.FindElement(By.Name("P_PortNbr")).Clear();
            driver.FindElement(By.Name("P_PortNbr")).SendKeys("502");
            driver.FindElement(By.Name("P_DevAddr")).Clear();
            driver.FindElement(By.Name("P_DevAddr")).SendKeys("1");
            driver.FindElement(By.Name("submit")).Click();

            driver.Navigate().Refresh();
        }

        private void CreateModbusTag()
        {
            long lErrorCode = 0;

            if (lErrorCode == 0)
            {
                driver.SwitchTo().ParentFrame();
                driver.SwitchTo().Frame("leftFrame");
                driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/device/devPg.asp') and contains(@href, 'dname=ModSim')]")).Click();
            }

            if (lErrorCode == 0)
            {
                driver.SwitchTo().ParentFrame();
                driver.SwitchTo().Frame("rightFrame");
                driver.FindElement(By.XPath("//a[2]/font/b")).Click();
            }
            
            // Create AI tag
            EventLog.AddLog("Create AI tags start...");
            CreateModbusAITags();

            // Create AO tag
            EventLog.AddLog("Create AO tags start...");
            CreateModbusAOTags();

            // Create DI tag
            EventLog.AddLog("Create DI tags start...");
            CreateModbusDITags();

            // Create DO tag
            EventLog.AddLog("Create DO tags start...");
            CreateModbusDOTags();

            /*
            // Create TEXT tag
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByValue("TEXT");
            driver.FindElement(By.Name("TagName")).Clear();
            driver.FindElement(By.Name("TagName")).SendKeys("AT_TX0001");
            driver.FindElement(By.Name("TextLen")).Clear();
            driver.FindElement(By.Name("TextLen")).SendKeys("100");
            driver.FindElement(By.Name("Value")).Clear();
            driver.FindElement(By.Name("Value")).SendKeys("for Auto Test");
            driver.FindElement(By.Name("Submit")).Click();
            */
        }

        private void CreateModbusAITags()
        {
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByText("AI");
            Thread.Sleep(1500);
            SetupBasicAnalogTagConfig();
            for (int i = 1; i <= 5; i++)
            {
                try
                {
                    driver.FindElement(By.Name("TagName")).Clear();
                    driver.FindElement(By.Name("TagName")).SendKeys("AT_AI0" + i.ToString("000"));
                    driver.FindElement(By.Name("Submit")).Click();
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("CreateModbusAITags error: " + ex.ToString());
                    bPartResult = false;
                }

            }
        }

        private void CreateModbusAOTags()
        {
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByText("AO");
            Thread.Sleep(1500);
            SetupBasicAnalogTagConfig();
            for (int i = 1; i <= 5; i++)
            {
                try
                {
                    driver.FindElement(By.Name("TagName")).Clear();
                    driver.FindElement(By.Name("TagName")).SendKeys("AT_AO0" + i.ToString("000"));
                    driver.FindElement(By.Name("Submit")).Click();
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("CreateModbusAOTags error: " + ex.ToString());
                    bPartResult = false;
                }
            }
        }

        private void CreateModbusDITags()
        {
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByText("DI");
            Thread.Sleep(1500);
            SetupBasicDigitalTagConfig();
            for (int i = 1; i <= 5; i++)
            {
                try
                {
                    driver.FindElement(By.Name("TagName")).Clear();
                    driver.FindElement(By.Name("TagName")).SendKeys("AT_DI0" + i.ToString("000"));
                    driver.FindElement(By.Name("Submit")).Click();
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("CreateModbusDITags error: " + ex.ToString());
                    bPartResult = false;
                }
            }
        }

        private void CreateModbusDOTags()
        {
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByText("DO");
            Thread.Sleep(1500);
            SetupBasicDigitalTagConfig();
            for (int i = 1; i <= 5; i++)
            {
                try
                {
                    driver.FindElement(By.Name("TagName")).Clear();
                    driver.FindElement(By.Name("TagName")).SendKeys("AT_DO0" + i.ToString("000"));
                    driver.FindElement(By.Name("Submit")).Click();
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("CreateModbusDOTags error: " + ex.ToString());
                    bPartResult = false;
                }
            }
        }

        private void SetupBasicAnalogTagConfig()
        {
            driver.FindElement(By.Name("Datalog")).Click();
            driver.FindElement(By.Name("DataLogDB")).Clear();
            driver.FindElement(By.Name("DataLogDB")).SendKeys("0");
            driver.FindElement(By.Name("ChangeLog")).Click();
            driver.FindElement(By.Name("SpanHigh")).Clear();
            driver.FindElement(By.Name("SpanHigh")).SendKeys("1000");
            driver.FindElement(By.Name("OutputHigh")).Clear();
            driver.FindElement(By.Name("OutputHigh")).SendKeys("1000");
            new SelectElement(driver.FindElement(By.Name("ReservedInt1"))).SelectByText("2");
            driver.FindElement(By.XPath("(//input[@name='LogTmRadio'])[2]")).Click();
        }

        private void SetupBasicDigitalTagConfig()
        {
            driver.FindElement(By.Name("Datalog")).Click();
            driver.FindElement(By.Name("DataLogDB")).Clear();
            driver.FindElement(By.Name("DataLogDB")).SendKeys("0");
            driver.FindElement(By.Name("ChangeLog")).Click();
            driver.FindElement(By.Name("ReservedInt1")).Click();
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
