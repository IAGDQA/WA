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

namespace CreateCalcTags
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
        string sTestItemName = "CreateCalcTags";
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

            //Create Calculate Tags test
            if (bPartResult == true)
            {
                EventLog.AddLog("Create Calculate Tags");
                sw.Reset(); sw.Start();
                try
                {
                    driver.SwitchTo().Frame("leftFrame");
                    if (wcf.IsTestElementPresent(driver, "XPath", "//a[contains(@href, '/broadWeb/bwMainRight.asp?pos=CalcList') and contains(@href, 'name=TestSCADA')]"))
                    {
                        driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwMainRight.asp?pos=CalcList') and contains(@href, 'name=TestSCADA')]")).Click();
                        Thread.Sleep(1500);
                        EventLog.AddLog("Delete the Calculate tag");
                        driver.SwitchTo().ParentFrame();
                        driver.SwitchTo().Frame("rightFrame");
                        driver.FindElement(By.Id("chk_selall")).Click();
                        Thread.Sleep(1000);
                        driver.FindElement(By.XPath("//a[2]/font/b")).Click();  // delete
                        Thread.Sleep(1000);
                        driver.SwitchTo().Alert().Accept();

                        Thread.Sleep(3000);
                        driver.Navigate().Refresh();
                        Thread.Sleep(3000);

                        // click 'TestSCADA' link at left frame 
                        driver.SwitchTo().ParentFrame();
                        driver.SwitchTo().Frame("leftFrame");
                        driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwMainRight.asp') and contains(@href, 'name=TestSCADA')]")).Click();

                    }
                    driver.SwitchTo().ParentFrame();

                    EventLog.AddLog("Create Calculate Tags...");
                    CreateCalculateTag();
                }
                catch (Exception ex)
                {
                    EventLog.AddLog(@"Error occurred create Calculate Tags: " + ex.ToString());
                    bPartResult = false;
                }
                sw.Stop();
                PrintStep("Create", "Create Calculate Tags", bPartResult, "None", sw.Elapsed.TotalMilliseconds.ToString());

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

        private void CreateCalculateTag()
        {
            driver.SwitchTo().Frame("rightFrame");
            //driver.FindElement(By.XPath("//a[5]/font/b")).Click();
            driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/bwMainRight.asp?pos=CalcList')]")).Click();
            driver.FindElement(By.XPath("//a[contains(@href, '/broadWeb/tag/tagPg.asp') and contains(@href, 'dt=CalcType')]")).Click();

            ////*Avgerage AI*//
            //EventLog.AddLog("CreateCal_Average_AI_Tag");
            //CreateCal_Average_AI_Tag();
            //PrintStep("CreateCal_Average_AI_Tag");

            ////*Avgerage AO*//
            //EventLog.AddLog("CreateCal_Average_AO_Tag");
            //CreateCal_Average_AO_Tag();
            //PrintStep("CreateCal_Average_AO_Tag");

            ////*Avgerage ConAna*//
            //EventLog.AddLog("CreateCal_Average_ConAna_Tag");
            //CreateCal_Average_ConAna_Tag();
            //PrintStep("CreateCal_Average_ConAna_Tag");

            ////*Avgerage System*//
            //EventLog.AddLog("CreateCal_Average_System_Tag");
            //CreateCal_Average_System_Tag();
            //PrintStep("CreateCal_Average_System_Tag");

            ////*Avgerage OPCDA*//
            //EventLog.AddLog("CreateCal_Average_OPCDA_Tag");
            //CreateCal_Average_OPCDA_Tag();
            //PrintStep("CreateCal_Average_OPCDA_Tag");

            ////*Avgerage OPCUA*//
            //EventLog.AddLog("CreateCal_Average_OPCUA_Tag");
            //CreateCal_Average_OPCUA_Tag();
            //PrintStep("CreateCal_Average_OPCUA_Tag");

            ////*AND DI*//
            //EventLog.AddLog("CreateCal_AND_DI_Tag");
            //CreateCal_AND_DI_Tag();
            //PrintStep("CreateCal_AND_DI_Tag");

            ////*AND DO*//
            //EventLog.AddLog("CreateCal_AND_DO_Tag");
            //CreateCal_AND_DO_Tag();
            //PrintStep("CreateCal_AND_DO_Tag");

            ////*AND ConDis*//
            //EventLog.AddLog("CreateCal_AND_ConDis_Tag");
            //CreateCal_AND_ConDis_Tag();
            //PrintStep("CreateCal_AND_ConDis_Tag");

            //// Analog tag /////
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByText("CalcAna");
            string[] CalcTagName_Ana = { "System", "ConAna", "ModBusAI", "ModBusAO", "OPCDA", "OPCUA", "Acc" };
            string[] Formula_Ana = { "MOD(A, 10)", "MOD(A, 10)", "MOD(A, 10)", "MOD(A, 10)", "MOD(A, 10)", "MOD(A, 10)", "MOD(A, 10)" };
            string[] VariableA_Ana = { "Calc_AvgSystemAll", "Calc_AvgConAnaAll", "Calc_AvgAIAll", "Calc_AvgAOAll", "Calc_AvgOPCDAAll", "Calc_AvgOPCUAAll", "Acc_0125" };
            Thread.Sleep(500);
            SetupBasicAnalogTagConfig();
            for (int i = 0; i < CalcTagName_Ana.Length; i++)
            {
                try
                {
                    driver.FindElement(By.Name("TagName")).Clear();
                    driver.FindElement(By.Name("TagName")).SendKeys("Calc_" + CalcTagName_Ana[i]);
                    driver.FindElement(By.Name("Formula")).Clear();
                    driver.FindElement(By.Name("Formula")).SendKeys(Formula_Ana[i] + " + " + (i * 10));
                    driver.FindElement(By.Name("A")).Clear();
                    driver.FindElement(By.Name("A")).SendKeys(VariableA_Ana[i]);
                    driver.FindElement(By.Name("Submit")).Click();
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create Analog CalcTags error: " + ex.ToString());
                    bPartResult = false;
                }
            }

            //// Discrete tag /////
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByText("CalcDis");
            string[] CalcTagName_Dis = { "ConDis", "ModBusDI", "ModBusDO"};
            string[] Formula_Dis = { "A*10","A*10","A*10",};
            string[] VariableA_Dis = { "Calc_ANDConDisAll", "Calc_ANDDIAll", "Calc_ANDDOAll" };
            Thread.Sleep(500);
            SetupBasicDiscreteTagConfig();
            for (int i = 0; i < CalcTagName_Dis.Length; i++)
            {
                try
                {
                    driver.FindElement(By.Name("TagName")).Clear();
                    driver.FindElement(By.Name("TagName")).SendKeys("Calc_" + CalcTagName_Dis[i]);
                    driver.FindElement(By.Name("Formula")).Clear();
                    driver.FindElement(By.Name("Formula")).SendKeys(Formula_Dis[i] + " + " + ((i + CalcTagName_Ana.Length) * 10));
                    driver.FindElement(By.Name("A")).Clear();
                    driver.FindElement(By.Name("A")).SendKeys(VariableA_Dis[i]);
                    driver.FindElement(By.Name("Submit")).Click();
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create Discrete CalcTags error: " + ex.ToString());
                    bPartResult = false;
                }
            }
        }

        private void CreateCal_Average_AI_Tag()
        {
            //// AvgAI020~AvgAI250 ////
            int iTagCount = 0;
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByText("CalcAna");
            string[] CalcTagName_Ana = { "AvgAI020", "AvgAI040", "AvgAI060", "AvgAI080", "AvgAI100", "AvgAI120", "AvgAI140",
                                         "AvgAI160", "AvgAI180", "AvgAI200", "AvgAI220", "AvgAI240", "AvgAI250"};
            Thread.Sleep(500);
            SetupBasicAnalogTagConfig();
            
            for (int i = 0; i < CalcTagName_Ana.Length; i++)
            {
                try
                {
                    driver.FindElement(By.Name("TagName")).Clear();
                    driver.FindElement(By.Name("TagName")).SendKeys("Calc_" + CalcTagName_Ana[i]);

                    driver.FindElement(By.Name("Formula")).Clear();
                    if (CalcTagName_Ana[i] == "AvgAI250")
                        driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J)/10");
                    else
                        driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J+K+L+M+N+O+P+Q+R+S+T)/20");

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Ana[i] == "AvgAI250")
                        {
                            if (j <= 10)
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("AT_AI0" + (iTagCount + j).ToString("000"));
                            }
                            else
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("");
                            }
                        }
                        else
                        {
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("AT_AI0" + (iTagCount + j).ToString("000"));
                        }

                        if (j == 20)
                            driver.FindElement(By.Name("Submit")).Click();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create AvgAI020~AvgAI250 CalcTags error: " + ex.ToString());
                    bPartResult = false;
                }

                iTagCount += 20;
            }

            //// AvgAIAll ////
            driver.FindElement(By.Name("TagName")).Clear();
            driver.FindElement(By.Name("TagName")).SendKeys("Calc_AvgAIAll");
            driver.FindElement(By.Name("Formula")).Clear();
            driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J+K+L+M)/13");
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Ana.Length)
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();    //ASCII 65 = "A"
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("Calc_" + CalcTagName_Ana[n - 1]);
                }
                else
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("");
                }
                if (n == 20)
                    driver.FindElement(By.Name("Submit")).Click();
            }
        }

        private void CreateCal_Average_AO_Tag()
        {
            //// AvgAO020~AvgAO250 ////
            int iTagCount = 0;
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByText("CalcAna");
            string[] CalcTagName_Ana = { "AvgAO020", "AvgAO040", "AvgAO060", "AvgAO080", "AvgAO100", "AvgAO120", "AvgAO140",
                                         "AvgAO160", "AvgAO180", "AvgAO200", "AvgAO220", "AvgAO240", "AvgAO250"};
            Thread.Sleep(500);
            SetupBasicAnalogTagConfig();
            for (int i = 0; i < CalcTagName_Ana.Length; i++)
            {
                try
                {
                    driver.FindElement(By.Name("TagName")).Clear();
                    driver.FindElement(By.Name("TagName")).SendKeys("Calc_" + CalcTagName_Ana[i]);

                    driver.FindElement(By.Name("Formula")).Clear();
                    if (CalcTagName_Ana[i] == "AvgAO250")
                        driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J)/10");
                    else
                        driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J+K+L+M+N+O+P+Q+R+S+T)/20");

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Ana[i] == "AvgAO250")
                        {
                            if (j <= 10)
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("AT_AO0" + (iTagCount + j).ToString("000"));
                            }
                            else
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("");
                            }
                        }
                        else
                        {
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("AT_AO0" + (iTagCount + j).ToString("000"));
                        }

                        if (j == 20)
                            driver.FindElement(By.Name("Submit")).Click();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create AvgAO020~AvgAO250 CalcTags error: " + ex.ToString());
                    bPartResult = false;
                }

                iTagCount += 20;
            }

            //// AvgAOAll ////
            driver.FindElement(By.Name("TagName")).Clear();
            driver.FindElement(By.Name("TagName")).SendKeys("Calc_AvgAOAll");
            driver.FindElement(By.Name("Formula")).Clear();
            driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J+K+L+M)/13");
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Ana.Length)
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();    //ASCII 65 = "A"
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("Calc_" + CalcTagName_Ana[n - 1]);
                }
                else
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("");
                }
                if (n == 20)
                    driver.FindElement(By.Name("Submit")).Click();
            }
        }

        private void CreateCal_Average_ConAna_Tag()
        {
            //// AvgConAna020~AvgConAna250 ////
            int iTagCount = 0;
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByText("CalcAna");
            string[] CalcTagName_Ana = { "AvgConAna020", "AvgConAna040", "AvgConAna060", "AvgConAna080", "AvgConAna100", "AvgConAna120", "AvgConAna140",
                                         "AvgConAna160", "AvgConAna180", "AvgConAna200", "AvgConAna220", "AvgConAna240", "AvgConAna250"};
            Thread.Sleep(500);
            SetupBasicAnalogTagConfig();
            for (int i = 0; i < CalcTagName_Ana.Length; i++)
            {
                try
                {
                    driver.FindElement(By.Name("TagName")).Clear();
                    driver.FindElement(By.Name("TagName")).SendKeys("Calc_" + CalcTagName_Ana[i]);

                    driver.FindElement(By.Name("Formula")).Clear();
                    if (CalcTagName_Ana[i] == "AvgConAna250")
                        driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J)/10");
                    else
                        driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J+K+L+M+N+O+P+Q+R+S+T)/20");

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Ana[i] == "AvgConAna250")
                        {
                            if (j <= 10)
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("ConAna_" + (iTagCount + j).ToString("0000"));
                            }
                            else
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("");
                            }
                        }
                        else
                        {
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("ConAna_" + (iTagCount + j).ToString("0000"));
                        }

                        if (j == 20)
                            driver.FindElement(By.Name("Submit")).Click();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create CalcTags error: " + ex.ToString());
                    bPartResult = false;
                }

                iTagCount += 20;
            }

            //// AvgConAnaAll ////
            driver.FindElement(By.Name("TagName")).Clear();
            driver.FindElement(By.Name("TagName")).SendKeys("Calc_AvgConAnaAll");
            driver.FindElement(By.Name("Formula")).Clear();
            driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J+K+L+M)/13");
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Ana.Length)
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();    //ASCII 65 = "A"
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("Calc_" + CalcTagName_Ana[n - 1]);
                }
                else
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("");
                }
                if (n == 20)
                    driver.FindElement(By.Name("Submit")).Click();
            }
        }

        private void CreateCal_Average_System_Tag()
        {
            //// AvgSystem020~AvgSystem250 ////
            int iTagCount = 0;
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByText("CalcAna");
            string[] CalcTagName_Ana = { "AvgSystem020", "AvgSystem040", "AvgSystem060", "AvgSystem080", "AvgSystem100", "AvgSystem120", "AvgSystem140",
                                         "AvgSystem160", "AvgSystem180", "AvgSystem200", "AvgSystem220", "AvgSystem240", "AvgSystem250"};
            Thread.Sleep(500);
            SetupBasicAnalogTagConfig();
            for (int i = 0; i < CalcTagName_Ana.Length; i++)
            {
                try
                {
                    driver.FindElement(By.Name("TagName")).Clear();
                    driver.FindElement(By.Name("TagName")).SendKeys("Calc_" + CalcTagName_Ana[i]);

                    driver.FindElement(By.Name("Formula")).Clear();
                    if (CalcTagName_Ana[i] == "AvgSystem250")
                        driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J)/10");
                    else
                        driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J+K+L+M+N+O+P+Q+R+S+T)/20");

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Ana[i] == "AvgSystem250")
                        {
                            if (j <= 10)
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("SystemSec_" + (iTagCount + j).ToString("0000"));
                            }
                            else
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("");
                            }
                        }
                        else
                        {
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("SystemSec_" + (iTagCount + j).ToString("0000"));
                        }

                        if (j == 20)
                            driver.FindElement(By.Name("Submit")).Click();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create CalcTags error: " + ex.ToString());
                    bPartResult = false;
                }

                iTagCount += 20;
            }

            //// AvgSystemAll ////
            driver.FindElement(By.Name("TagName")).Clear();
            driver.FindElement(By.Name("TagName")).SendKeys("Calc_AvgSystemAll");
            driver.FindElement(By.Name("Formula")).Clear();
            driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J+K+L+M)/13");
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Ana.Length)
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();    //ASCII 65 = "A"
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("Calc_" + CalcTagName_Ana[n - 1]);
                }
                else
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("");
                }
                if (n == 20)
                    driver.FindElement(By.Name("Submit")).Click();
            }
        }

        private void CreateCal_Average_OPCDA_Tag()
        {
            //// AvgOPCDA020~AvgOPCDA250 ////
            int iTagCount = 0;
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByText("CalcAna");
            string[] CalcTagName_Ana = { "AvgOPCDA020", "AvgOPCDA040", "AvgOPCDA060", "AvgOPCDA080", "AvgOPCDA100", "AvgOPCDA120", "AvgOPCDA140",
                                         "AvgOPCDA160", "AvgOPCDA180", "AvgOPCDA200", "AvgOPCDA220", "AvgOPCDA240", "AvgOPCDA250"};
            Thread.Sleep(500);
            SetupBasicAnalogTagConfig();
            for (int i = 0; i < CalcTagName_Ana.Length; i++)
            {
                try
                {
                    driver.FindElement(By.Name("TagName")).Clear();
                    driver.FindElement(By.Name("TagName")).SendKeys("Calc_" + CalcTagName_Ana[i]);

                    driver.FindElement(By.Name("Formula")).Clear();
                    if (CalcTagName_Ana[i] == "AvgOPCDA250")
                        driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J)/10");
                    else
                        driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J+K+L+M+N+O+P+Q+R+S+T)/20");

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Ana[i] == "AvgOPCDA250")
                        {
                            if (j <= 10)
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("OPCDA_" + (iTagCount + j).ToString("0000"));
                            }
                            else
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("");
                            }
                        }
                        else
                        {
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("OPCDA_" + (iTagCount + j).ToString("0000"));
                        }

                        if (j == 20)
                            driver.FindElement(By.Name("Submit")).Click();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create CalcTags error: " + ex.ToString());
                    bPartResult = false;
                }

                iTagCount += 20;
            }

            //// AvgOPCDAAll ////
            driver.FindElement(By.Name("TagName")).Clear();
            driver.FindElement(By.Name("TagName")).SendKeys("Calc_AvgOPCDAAll");
            driver.FindElement(By.Name("Formula")).Clear();
            driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J+K+L+M)/13");
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Ana.Length)
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();    //ASCII 65 = "A"
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("Calc_" + CalcTagName_Ana[n - 1]);
                }
                else
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("");
                }
                if (n == 20)
                    driver.FindElement(By.Name("Submit")).Click();
            }
        }

        private void CreateCal_Average_OPCUA_Tag()
        {
            //// AvgOPCUA020~AvgOPCUA250 ////
            int iTagCount = 0;
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByText("CalcAna");
            string[] CalcTagName_Ana = { "AvgOPCUA020", "AvgOPCUA040", "AvgOPCUA060", "AvgOPCUA080", "AvgOPCUA100", "AvgOPCUA120", "AvgOPCUA140",
                                         "AvgOPCUA160", "AvgOPCUA180", "AvgOPCUA200", "AvgOPCUA220", "AvgOPCUA240", "AvgOPCUA250"};
            Thread.Sleep(500);
            SetupBasicAnalogTagConfig();
            for (int i = 0; i < CalcTagName_Ana.Length; i++)
            {
                try
                {
                    driver.FindElement(By.Name("TagName")).Clear();
                    driver.FindElement(By.Name("TagName")).SendKeys("Calc_" + CalcTagName_Ana[i]);

                    driver.FindElement(By.Name("Formula")).Clear();
                    if (CalcTagName_Ana[i] == "AvgOPCUA250")
                        driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J)/10");
                    else
                        driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J+K+L+M+N+O+P+Q+R+S+T)/20");

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Ana[i] == "AvgOPCUA250")
                        {
                            if (j <= 10)
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("OPCUA_" + (iTagCount + j).ToString("0000"));
                            }
                            else
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("");
                            }
                        }
                        else
                        {
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("OPCUA_" + (iTagCount + j).ToString("0000"));
                        }

                        if (j == 20)
                            driver.FindElement(By.Name("Submit")).Click();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create CalcTags error: " + ex.ToString());
                    bPartResult = false;
                }

                iTagCount += 20;
            }

            //// AvgOPCUAAll ////
            driver.FindElement(By.Name("TagName")).Clear();
            driver.FindElement(By.Name("TagName")).SendKeys("Calc_AvgOPCUAAll");
            driver.FindElement(By.Name("Formula")).Clear();
            driver.FindElement(By.Name("Formula")).SendKeys("(A+B+C+D+E+F+G+H+I+J+K+L+M)/13");
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Ana.Length)
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();    //ASCII 65 = "A"
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("Calc_" + CalcTagName_Ana[n - 1]);
                }
                else
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("");
                }
                if (n == 20)
                    driver.FindElement(By.Name("Submit")).Click();
            }
        }

        private void CreateCal_AND_DI_Tag()
        {
            //// ANDDI020~ANDDI250 ////
            int iTagCount = 0;
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByText("CalcDis");
            string[] CalcTagName_Dis = { "ANDDI020", "ANDDI040", "ANDDI060", "ANDDI080", "ANDDI100", "ANDDI120", "ANDDI140",
                                         "ANDDI160", "ANDDI180", "ANDDI200", "ANDDI220", "ANDDI240", "ANDDI250"};
            Thread.Sleep(500);
            SetupBasicDiscreteTagConfig();

            for (int i = 0; i < CalcTagName_Dis.Length; i++)
            {
                try
                {
                    driver.FindElement(By.Name("TagName")).Clear();
                    driver.FindElement(By.Name("TagName")).SendKeys("Calc_" + CalcTagName_Dis[i]);

                    driver.FindElement(By.Name("Formula")).Clear();
                    if (CalcTagName_Dis[i] == "ANDDI250")
                        driver.FindElement(By.Name("Formula")).SendKeys("A&&B&&C&&D&&E&&F&&G&&H&&I&&J");
                    else
                        driver.FindElement(By.Name("Formula")).SendKeys("A&&B&&C&&D&&E&&F&&G&&H&&I&&J&&K&&L&&M&&N&&O&&P&&Q&&R&&S&&T");

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Dis[i] == "ANDDI250")
                        {
                            if (j <= 10)
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("AT_DI0" + (iTagCount + j).ToString("000"));
                            }
                            else
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("");
                            }
                        }
                        else
                        {
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("AT_DI0" + (iTagCount + j).ToString("000"));
                        }

                        if (j == 20)
                            driver.FindElement(By.Name("Submit")).Click();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create CalcTags error: " + ex.ToString());
                    bPartResult = false;
                }

                iTagCount += 20;
            }

            //// ANDDIAll ////
            driver.FindElement(By.Name("TagName")).Clear();
            driver.FindElement(By.Name("TagName")).SendKeys("Calc_ANDDIAll");
            driver.FindElement(By.Name("Formula")).Clear();
            driver.FindElement(By.Name("Formula")).SendKeys("A&&B&&C&&D&&E&&F&&G&&H&&I&&J&&K&&L&&M");
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Dis.Length)
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();    //ASCII 65 = "A"
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("Calc_" + CalcTagName_Dis[n - 1]);
                }
                else
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("");
                }
                if (n == 20)
                    driver.FindElement(By.Name("Submit")).Click();
            }
        }

        private void CreateCal_AND_DO_Tag()
        {
            //// ANDDO020~ANDDO250 ////
            int iTagCount = 0;
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByText("CalcDis");
            string[] CalcTagName_Dis = { "ANDDO020", "ANDDO040", "ANDDO060", "ANDDO080", "ANDDO100", "ANDDO120", "ANDDO140",
                                         "ANDDO160", "ANDDO180", "ANDDO200", "ANDDO220", "ANDDO240", "ANDDO250"};
            Thread.Sleep(500);
            SetupBasicDiscreteTagConfig();

            for (int i = 0; i < CalcTagName_Dis.Length; i++)
            {
                try
                {
                    driver.FindElement(By.Name("TagName")).Clear();
                    driver.FindElement(By.Name("TagName")).SendKeys("Calc_" + CalcTagName_Dis[i]);

                    driver.FindElement(By.Name("Formula")).Clear();
                    if (CalcTagName_Dis[i] == "ANDDO250")
                        driver.FindElement(By.Name("Formula")).SendKeys("A&&B&&C&&D&&E&&F&&G&&H&&I&&J");
                    else
                        driver.FindElement(By.Name("Formula")).SendKeys("A&&B&&C&&D&&E&&F&&G&&H&&I&&J&&K&&L&&M&&N&&O&&P&&Q&&R&&S&&T");

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Dis[i] == "ANDDO250")
                        {
                            if (j <= 10)
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("AT_DO0" + (iTagCount + j).ToString("000"));
                            }
                            else
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("");
                            }
                        }
                        else
                        {
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("AT_DO0" + (iTagCount + j).ToString("000"));
                        }

                        if (j == 20)
                            driver.FindElement(By.Name("Submit")).Click();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create CalcTags error: " + ex.ToString());
                    bPartResult = false;
                }

                iTagCount += 20;
            }

            //// ANDDOAll ////
            driver.FindElement(By.Name("TagName")).Clear();
            driver.FindElement(By.Name("TagName")).SendKeys("Calc_ANDDOAll");
            driver.FindElement(By.Name("Formula")).Clear();
            driver.FindElement(By.Name("Formula")).SendKeys("A&&B&&C&&D&&E&&F&&G&&H&&I&&J&&K&&L&&M");
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Dis.Length)
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();    //ASCII 65 = "A"
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("Calc_" + CalcTagName_Dis[n - 1]);
                }
                else
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("");
                }
                if (n == 20)
                    driver.FindElement(By.Name("Submit")).Click();
            }
        }

        private void CreateCal_AND_ConDis_Tag()
        {
            //// ANDConDis020~ANDConDis250 ////
            int iTagCount = 0;
            new SelectElement(driver.FindElement(By.Name("ParaName"))).SelectByText("CalcDis");
            string[] CalcTagName_Dis = { "ANDConDis020", "ANDConDis040", "ANDConDis060", "ANDConDis080", "ANDConDis100", "ANDConDis120", "ANDConDis140",
                                         "ANDConDis160", "ANDConDis180", "ANDConDis200", "ANDConDis220", "ANDConDis240", "ANDConDis250"};
            Thread.Sleep(500);
            SetupBasicDiscreteTagConfig();

            for (int i = 0; i < CalcTagName_Dis.Length; i++)
            {
                try
                {
                    driver.FindElement(By.Name("TagName")).Clear();
                    driver.FindElement(By.Name("TagName")).SendKeys("Calc_" + CalcTagName_Dis[i]);

                    driver.FindElement(By.Name("Formula")).Clear();
                    if (CalcTagName_Dis[i] == "ANDConDis250")
                        driver.FindElement(By.Name("Formula")).SendKeys("A&&B&&C&&D&&E&&F&&G&&H&&I&&J");
                    else
                        driver.FindElement(By.Name("Formula")).SendKeys("A&&B&&C&&D&&E&&F&&G&&H&&I&&J&&K&&L&&M&&N&&O&&P&&Q&&R&&S&&T");

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Dis[i] == "ANDConDis250")
                        {
                            if (j <= 10)
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("ConDis_" + (iTagCount + j).ToString("0000"));
                            }
                            else
                            {
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();
                                driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("");
                            }
                        }
                        else
                        {
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).Clear();    //ASCII 65 = "A"
                            driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(j + 64)))).SendKeys("ConDis_" + (iTagCount + j).ToString("0000"));
                        }

                        if (j == 20)
                            driver.FindElement(By.Name("Submit")).Click();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create CalcTags error: " + ex.ToString());
                    bPartResult = false;
                }

                iTagCount += 20;
            }

            //// ANDConDisAll ////
            driver.FindElement(By.Name("TagName")).Clear();
            driver.FindElement(By.Name("TagName")).SendKeys("Calc_ANDConDisAll");
            driver.FindElement(By.Name("Formula")).Clear();
            driver.FindElement(By.Name("Formula")).SendKeys("A&&B&&C&&D&&E&&F&&G&&H&&I&&J&&K&&L&&M");
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Dis.Length)
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();    //ASCII 65 = "A"
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("Calc_" + CalcTagName_Dis[n - 1]);
                }
                else
                {
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).Clear();
                    driver.FindElement(By.Name(System.Convert.ToString(Convert.ToChar(n + 64)))).SendKeys("");
                }
                if (n == 20)
                    driver.FindElement(By.Name("Submit")).Click();
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

        private void SetupBasicDiscreteTagConfig()
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
