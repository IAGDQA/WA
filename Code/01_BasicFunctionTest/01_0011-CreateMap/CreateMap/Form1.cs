using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ThirdPartyToolControl;
using iATester;
using CommonFunction;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;       // for SelectElement use
using System.Diagnostics;

namespace CreateMap
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
        string sTestItemName = "CreateMap";
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

            //Create Map
            if (bPartResult == true)
            {
                EventLog.AddLog("Create Map");
                sw.Reset(); sw.Start();
                try
                {
                    CreateMap(sTestLogFolder, sLanguage);
                }
                catch (Exception ex)
                {
                    EventLog.AddLog(@"Error occurred Create Map: " + ex.ToString());
                    bPartResult = false;
                }
                sw.Stop();
                PrintStep("Create", "Create Map", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());

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

        private void CreateMap(string sTestLogFolder, string slanguage)
        {
            driver.SwitchTo().Frame("rightFrame");

            if (slanguage == "CHS")     // fuck china special case..
            {
                driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bmap/bmapcreate.asp?')]")).Click();
                System.Threading.Thread.Sleep(2000);
                if (!isAlertPresent())
                {
                    EventLog.AddLog("Click 'New Google Map' test");
                    driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/gmap/gmapcreate.asp?')]")).Click();
                    EventLog.PrintScreen("CreateMapTest_GoogleMap");
                }
                else
                {
                    bPartResult = false;
                }
                //PrintStep("Google Map click test");
            }
            else
            {
                driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/gmap/gmapcreate.asp')]")).Click();
                System.Threading.Thread.Sleep(2000);
                if (!isAlertPresent())
                {
                    //TestGoogleMap/BaiduMap
                    EventLog.AddLog("Click 'New Baidu Map' test");
                    driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bmap/bmapcreate.asp?')]")).Click();
                    EventLog.PrintScreen("CreateMapTest_BiaduMap");

                    EventLog.AddLog("Click 'New Google Map' test");
                    driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/gmap/gmapcreate.asp?')]")).Click();
                    EventLog.PrintScreen("CreateMapTest_GoogleMap");
                    //PrintStep("Google&Baidu Map click test");
                }
                else
                {
                    bPartResult = false;
                }
            }

            if (bPartResult == true)
            {
                //Excel-In sample map
                EventLog.AddLog("Excel in sample map");
                driver.FindElement(By.XPath("//a[contains(@href, 'gmaptoJsPg1.asp?pos=import')]")).Click();
                string sCurrentFilePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetAssembly(this.GetType()).Location);
                string sourceSampleFile = sCurrentFilePath + "\\MapSample\\MapSample.xls";
                string destWApath = @"C:\Inetpub\wwwroot\broadweb\gmap\MapSample.xls";
                System.IO.File.Copy(sourceSampleFile, destWApath, true);
                driver.FindElement(By.Name("dataFileName")).Clear();
                driver.FindElement(By.Name("dataFileName")).SendKeys("MapSample");
                driver.FindElement(By.Name("submit")).Click();
                driver.FindElement(By.Name("act")).Click();
                System.Threading.Thread.Sleep(2000);
                EventLog.PrintScreen("CreateMapTest_Import_SampleMap");
                EventLog.AddLog("Excel In Map");

                //Options
                EventLog.AddLog("Options setting...");
                EventLog.AddLog("Marker title font set");
                driver.FindElement(By.XPath("(//a[contains(@href, '#')])[4]")).Click();
                System.Threading.Thread.Sleep(1000);
                driver.FindElement(By.XPath("(//input[@name='aa'])[2]")).Click();
                System.Threading.Thread.Sleep(500);
                driver.FindElement(By.XPath("(//input[@name='aa'])[3]")).Click();
                System.Threading.Thread.Sleep(500);
                driver.FindElement(By.XPath("(//input[@name='aa'])[1]")).Click();
                System.Threading.Thread.Sleep(500);

                driver.FindElement(By.XPath("(//input[@name='bb'])[2]")).Click();
                System.Threading.Thread.Sleep(500);
                driver.FindElement(By.XPath("(//input[@name='bb'])[1]")).Click();
                System.Threading.Thread.Sleep(500);

                driver.FindElement(By.XPath("//input[@id='cc']")).Click();
                driver.FindElement(By.XPath("//div[@id='fontpicker']/div")).Click(); //Font Family = "Microsoft YaHei"
                System.Threading.Thread.Sleep(500);
                driver.FindElement(By.XPath("//select[@id='dd']")).Click();
                driver.FindElement(By.XPath("//select[@id='dd']")).SendKeys("16");   //Font Size = 16
                System.Threading.Thread.Sleep(500);
                //api.ById("ee").Clear();
                //api.ById("ee").Enter("FF0000").Exe();   //Title Color = RED
                //System.Threading.Thread.Sleep(1000);
                EventLog.AddLog("Marker Title Font setting");

                EventLog.AddLog("Marker label font set");
                driver.FindElement(By.XPath("//input[@id='ff']")).Click();
                driver.FindElement(By.XPath("//div[@id='fontpicker']/div[10]")).Click(); //Font Family = "Impact"
                System.Threading.Thread.Sleep(500);
                driver.FindElement(By.XPath("//select[@id='gg']")).Click();
                driver.FindElement(By.XPath("//select[@id='gg']")).SendKeys("16");   //Font Size = 16
                System.Threading.Thread.Sleep(500);
                driver.FindElement(By.Id("hh")).Clear();
                driver.FindElement(By.Id("hh")).SendKeys("0000FF");   //Title Color = Bule
                System.Threading.Thread.Sleep(500);

                driver.FindElement(By.Id("ee")).Clear();
                driver.FindElement(By.Id("ee")).SendKeys("FF00EE");   //Title Color = Purple
                System.Threading.Thread.Sleep(500);

                driver.FindElement(By.XPath("//div[@id='opt']/div[27]/input")).Click();
                System.Threading.Thread.Sleep(500);
                EventLog.AddLog("Marker Label Font");

                //Save
                EventLog.AddLog("Save map");
                driver.FindElement(By.XPath("(//a[contains(@href, '#')])[2]")).Click();
                System.Threading.Thread.Sleep(1000);
                SendKeys.SendWait("{ENTER}");
                System.Threading.Thread.Sleep(1000);
                EventLog.AddLog("Save");
                EventLog.PrintScreen("CreateMapTest_ModifiedMap");
                System.Threading.Thread.Sleep(1000);

                //Excel-Out
                EventLog.AddLog("Excel out modified map");
                driver.FindElement(By.XPath("//a[contains(@href, 'gmaptoJsPg1.asp?pos=export')]")).Click();
                driver.FindElement(By.Name("chk")).Click();
                driver.FindElement(By.Name("dataFileName")).Clear();
                driver.FindElement(By.Name("dataFileName")).SendKeys("gmap_" + DateTime.Now.ToString("yyyyMMdd"));
                driver.FindElement(By.Name("submit")).Click();
                driver.FindElement(By.Name("act")).Click();
                EventLog.AddLog("Excel Out Map");

                try
                {
                    string sourceFile = @"C:\Inetpub\wwwroot\broadweb\gmap\gmap_" + DateTime.Now.ToString("yyyyMMdd") + ".xls";
                    string destFile = sTestLogFolder + "\\CreateMapTest_" + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".xls";
                    EventLog.AddLog("Copy export file form " + sourceFile + " to " + destFile);
                    System.IO.File.Copy(sourceFile, destFile, true);
                    System.Threading.Thread.Sleep(1000);
                }
                catch (Exception ex)
                {
                    EventLog.AddLog(ex.ToString());
                    bPartResult = false;
                }

                //Delete
                EventLog.AddLog("Delete map");
                driver.FindElement(By.XPath("(//a[contains(@href, '#')])[3]")).Click();
                System.Threading.Thread.Sleep(1000);
                driver.SwitchTo().Alert().Accept();
                System.Threading.Thread.Sleep(1000);
                EventLog.AddLog("Delete");
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

        private bool isAlertPresent()
        {
            try
            {
                EventLog.AddLog("Check if alert window pop up?");
                if (driver.SwitchTo().Alert() != null)
                {
                    EventLog.AddLog("Pop up alert window message: " + driver.SwitchTo().Alert().Text);
                    driver.SwitchTo().Alert().Accept();
                }
                return true;
            }
            catch (Exception ex)
            {
                EventLog.AddLog("isAlertPresent check: " + ex.ToString());
                return false;
            }
        }
    }
}
