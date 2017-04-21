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
using CommonFunction;

namespace CreateCalcTags
{
    public partial class Form1 : Form, iATester.iCom
    {
        IAdvSeleniumAPI api;
        cThirdPartyToolControl tpc = new cThirdPartyToolControl();
        cEventLog EventLog = new cEventLog();

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
            EventLog.AddLog("===Create calc tags start (by iATester)===");
            if (System.IO.File.Exists(sIniFilePath))    // 再load一次
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===Create calc tags end (by iATester)===");

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

            //Step1
            EventLog.AddLog("Create Calculate Tags...");
            ReturnSCADAPage();
            CreateCalculateTag();
            PrintStep("Create Calculate Tags");

            //driver.Close();
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

        private void CreateCalculateTag()
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//a[5]/font/b").Click();
            api.ByXpath("//a[contains(@href, '/broadWeb/tag/tagPg.asp') and contains(@href, 'dt=CalcType')]").Click();
            
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
            api.ByName("ParaName").SelectVal("CalcAna").Exe();
            string[] CalcTagName_Ana = { "System", "ConAna", "ModBusAI", "ModBusAO", "OPCDA", "OPCUA", "Acc" };
            string[] Formula_Ana = { "MOD(A, 10)", "MOD(A, 10)", "MOD(A, 10)", "MOD(A, 10)", "MOD(A, 10)", "MOD(A, 10)", "MOD(A, 10)" };
            string[] VariableA_Ana = { "Calc_AvgSystemAll", "Calc_AvgConAnaAll", "Calc_AvgAIAll", "Calc_AvgAOAll", "Calc_AvgOPCDAAll", "Calc_AvgOPCUAAll", "Acc_0125" };
            Thread.Sleep(500);
            SetupBasicAnalogTagConfig();
            for (int i = 0; i < CalcTagName_Ana.Length; i++)
            {
                try
                {
                    api.ByName("TagName").Clear();
                    api.ByName("TagName").Enter("Calc_" + CalcTagName_Ana[i]).Exe();
                    api.ByName("Formula").Clear();
                    api.ByName("Formula").Enter(Formula_Ana[i] + " + " + (i * 10)).Exe();
                    api.ByName("A").Clear();
                    api.ByName("A").Enter(VariableA_Ana[i]).Submit().Exe();
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create Analog CalcTags error: " + ex.ToString());
                    i--;
                }
            }

            //// Discrete tag /////
            api.ByName("ParaName").SelectVal("CalcDis").Exe();
            string[] CalcTagName_Dis = { "ConDis", "ModBusDI", "ModBusDO"};
            string[] Formula_Dis = { "A*10","A*10","A*10",};
            string[] VariableA_Dis = { "Calc_ANDConDisAll", "Calc_ANDDIAll", "Calc_ANDDOAll" };
            Thread.Sleep(500);
            SetupBasicDiscreteTagConfig();
            for (int i = 0; i < CalcTagName_Dis.Length; i++)
            {
                try
                {
                    api.ByName("TagName").Clear();
                    api.ByName("TagName").Enter("Calc_" + CalcTagName_Dis[i]).Exe();
                    api.ByName("Formula").Clear();
                    api.ByName("Formula").Enter(Formula_Dis[i] + " + " + ((i + CalcTagName_Ana.Length) * 10)).Exe();
                    api.ByName("A").Clear();
                    api.ByName("A").Enter(VariableA_Dis[i]).Submit().Exe();
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create Discrete CalcTags error: " + ex.ToString());
                    i--;
                }
            }
        }

        private void ReturnSCADAPage()
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMainRight.asp') and contains(@href, 'name=TestSCADA')]").Click();

        }

        private void CreateCal_Average_AI_Tag()
        {
            //// AvgAI020~AvgAI250 ////
            int iTagCount = 0;
            api.ByName("ParaName").SelectVal("CalcAna").Exe();
            string[] CalcTagName_Ana = { "AvgAI020", "AvgAI040", "AvgAI060", "AvgAI080", "AvgAI100", "AvgAI120", "AvgAI140",
                                         "AvgAI160", "AvgAI180", "AvgAI200", "AvgAI220", "AvgAI240", "AvgAI250"};
            Thread.Sleep(500);
            SetupBasicAnalogTagConfig();
            
            for (int i = 0; i < CalcTagName_Ana.Length; i++)
            {
                try
                {
                    api.ByName("TagName").Clear();
                    api.ByName("TagName").Enter("Calc_" + CalcTagName_Ana[i]).Exe();

                    api.ByName("Formula").Clear();
                    if (CalcTagName_Ana[i] == "AvgAI250")
                        api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J)/10").Exe();
                    else
                        api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J+K+L+M+N+O+P+Q+R+S+T)/20").Exe();

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Ana[i] == "AvgAI250")
                        {
                            if (j <= 10)
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("AT_AI0" + (iTagCount + j).ToString("000")).Exe();
                            }
                            else
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("").Exe();
                            }
                        }
                        else
                        {
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("AT_AI0" + (iTagCount + j).ToString("000")).Exe();
                        }

                        if (j == 20)
                            api.ByName("T").Enter("").Submit().Exe();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create AvgAI020~AvgAI250 CalcTags error: " + ex.ToString());
                    i--;
                }

                iTagCount += 20;
            }
            
            //// AvgAIAll ////
            api.ByName("TagName").Clear();
            api.ByName("TagName").Enter("Calc_AvgAIAll").Exe();
            api.ByName("Formula").Clear();
            api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J+K+L+M)/13").Exe();
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Ana.Length)
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();    //ASCII 65 = "A"
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("Calc_" + CalcTagName_Ana[n - 1]).Exe();
                }
                else
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("").Exe();
                }
                if (n == 20)
                    api.ByName("T").Enter("").Submit().Exe();
            }
        }

        private void CreateCal_Average_AO_Tag()
        {
            //// AvgAO020~AvgAO250 ////
            int iTagCount = 0;
            api.ByName("ParaName").SelectVal("CalcAna").Exe();
            string[] CalcTagName_Ana = { "AvgAO020", "AvgAO040", "AvgAO060", "AvgAO080", "AvgAO100", "AvgAO120", "AvgAO140",
                                         "AvgAO160", "AvgAO180", "AvgAO200", "AvgAO220", "AvgAO240", "AvgAO250"};
            Thread.Sleep(500);
            SetupBasicAnalogTagConfig();
            for (int i = 0; i < CalcTagName_Ana.Length; i++)
            {
                try
                {
                    api.ByName("TagName").Clear();
                    api.ByName("TagName").Enter("Calc_" + CalcTagName_Ana[i]).Exe();

                    api.ByName("Formula").Clear();
                    if (CalcTagName_Ana[i] == "AvgAO250")
                        api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J)/10").Exe();
                    else
                        api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J+K+L+M+N+O+P+Q+R+S+T)/20").Exe();

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Ana[i] == "AvgAO250")
                        {
                            if (j <= 10)
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("AT_AO0" + (iTagCount + j).ToString("000")).Exe();
                            }
                            else
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("").Exe();
                            }
                        }
                        else
                        {
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("AT_AO0" + (iTagCount + j).ToString("000")).Exe();
                        }

                        if (j == 20)
                            api.ByName("T").Enter("").Submit().Exe();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create AvgAO020~AvgAO250 CalcTags error: " + ex.ToString());
                    i--;
                }

                iTagCount += 20;
            }

            //// AvgAOAll ////
            api.ByName("TagName").Clear();
            api.ByName("TagName").Enter("Calc_AvgAOAll").Exe();
            api.ByName("Formula").Clear();
            api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J+K+L+M)/13").Exe();
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Ana.Length)
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();    //ASCII 65 = "A"
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("Calc_" + CalcTagName_Ana[n - 1]).Exe();
                }
                else
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("").Exe();
                }
                if (n == 20)
                    api.ByName("T").Enter("").Submit().Exe();
            }
        }

        private void CreateCal_Average_ConAna_Tag()
        {
            //// AvgConAna020~AvgConAna250 ////
            int iTagCount = 0;
            api.ByName("ParaName").SelectVal("CalcAna").Exe();
            string[] CalcTagName_Ana = { "AvgConAna020", "AvgConAna040", "AvgConAna060", "AvgConAna080", "AvgConAna100", "AvgConAna120", "AvgConAna140",
                                         "AvgConAna160", "AvgConAna180", "AvgConAna200", "AvgConAna220", "AvgConAna240", "AvgConAna250"};
            Thread.Sleep(500);
            SetupBasicAnalogTagConfig();
            for (int i = 0; i < CalcTagName_Ana.Length; i++)
            {
                try
                {
                    api.ByName("TagName").Clear();
                    api.ByName("TagName").Enter("Calc_" + CalcTagName_Ana[i]).Exe();

                    api.ByName("Formula").Clear();
                    if (CalcTagName_Ana[i] == "AvgConAna250")
                        api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J)/10").Exe();
                    else
                        api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J+K+L+M+N+O+P+Q+R+S+T)/20").Exe();

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Ana[i] == "AvgConAna250")
                        {
                            if (j <= 10)
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("ConAna_" + (iTagCount + j).ToString("0000")).Exe();
                            }
                            else
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("").Exe();
                            }
                        }
                        else
                        {
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("ConAna_" + (iTagCount + j).ToString("0000")).Exe();
                        }

                        if (j == 20)
                            api.ByName("T").Enter("").Submit().Exe();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create CalcTags error: " + ex.ToString());
                    i--;
                }

                iTagCount += 20;
            }

            //// AvgConAnaAll ////
            api.ByName("TagName").Clear();
            api.ByName("TagName").Enter("Calc_AvgConAnaAll").Exe();
            api.ByName("Formula").Clear();
            api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J+K+L+M)/13").Exe();
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Ana.Length)
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();    //ASCII 65 = "A"
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("Calc_" + CalcTagName_Ana[n - 1]).Exe();
                }
                else
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("").Exe();
                }
                if (n == 20)
                    api.ByName("T").Enter("").Submit().Exe();
            }
        }

        private void CreateCal_Average_System_Tag()
        {
            //// AvgSystem020~AvgSystem250 ////
            int iTagCount = 0;
            api.ByName("ParaName").SelectVal("CalcAna").Exe();
            string[] CalcTagName_Ana = { "AvgSystem020", "AvgSystem040", "AvgSystem060", "AvgSystem080", "AvgSystem100", "AvgSystem120", "AvgSystem140",
                                         "AvgSystem160", "AvgSystem180", "AvgSystem200", "AvgSystem220", "AvgSystem240", "AvgSystem250"};
            Thread.Sleep(500);
            SetupBasicAnalogTagConfig();
            for (int i = 0; i < CalcTagName_Ana.Length; i++)
            {
                try
                {
                    api.ByName("TagName").Clear();
                    api.ByName("TagName").Enter("Calc_" + CalcTagName_Ana[i]).Exe();

                    api.ByName("Formula").Clear();
                    if (CalcTagName_Ana[i] == "AvgSystem250")
                        api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J)/10").Exe();
                    else
                        api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J+K+L+M+N+O+P+Q+R+S+T)/20").Exe();

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Ana[i] == "AvgSystem250")
                        {
                            if (j <= 10)
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("SystemSec_" + (iTagCount + j).ToString("0000")).Exe();
                            }
                            else
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("").Exe();
                            }
                        }
                        else
                        {
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("SystemSec_" + (iTagCount + j).ToString("0000")).Exe();
                        }

                        if (j == 20)
                            api.ByName("T").Enter("").Submit().Exe();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create CalcTags error: " + ex.ToString());
                    i--;
                }

                iTagCount += 20;
            }

            //// AvgSystemAll ////
            api.ByName("TagName").Clear();
            api.ByName("TagName").Enter("Calc_AvgSystemAll").Exe();
            api.ByName("Formula").Clear();
            api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J+K+L+M)/13").Exe();
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Ana.Length)
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();    //ASCII 65 = "A"
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("Calc_" + CalcTagName_Ana[n - 1]).Exe();
                }
                else
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("").Exe();
                }
                if (n == 20)
                    api.ByName("T").Enter("").Submit().Exe();
            }
        }

        private void CreateCal_Average_OPCDA_Tag()
        {
            //// AvgOPCDA020~AvgOPCDA250 ////
            int iTagCount = 0;
            api.ByName("ParaName").SelectVal("CalcAna").Exe();
            string[] CalcTagName_Ana = { "AvgOPCDA020", "AvgOPCDA040", "AvgOPCDA060", "AvgOPCDA080", "AvgOPCDA100", "AvgOPCDA120", "AvgOPCDA140",
                                         "AvgOPCDA160", "AvgOPCDA180", "AvgOPCDA200", "AvgOPCDA220", "AvgOPCDA240", "AvgOPCDA250"};
            Thread.Sleep(500);
            SetupBasicAnalogTagConfig();
            for (int i = 0; i < CalcTagName_Ana.Length; i++)
            {
                try
                {
                    api.ByName("TagName").Clear();
                    api.ByName("TagName").Enter("Calc_" + CalcTagName_Ana[i]).Exe();

                    api.ByName("Formula").Clear();
                    if (CalcTagName_Ana[i] == "AvgOPCDA250")
                        api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J)/10").Exe();
                    else
                        api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J+K+L+M+N+O+P+Q+R+S+T)/20").Exe();

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Ana[i] == "AvgOPCDA250")
                        {
                            if (j <= 10)
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("OPCDA_" + (iTagCount + j).ToString("0000")).Exe();
                            }
                            else
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("").Exe();
                            }
                        }
                        else
                        {
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("OPCDA_" + (iTagCount + j).ToString("0000")).Exe();
                        }

                        if (j == 20)
                            api.ByName("T").Enter("").Submit().Exe();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create CalcTags error: " + ex.ToString());
                    i--;
                }

                iTagCount += 20;
            }

            //// AvgOPCDAAll ////
            api.ByName("TagName").Clear();
            api.ByName("TagName").Enter("Calc_AvgOPCDAAll").Exe();
            api.ByName("Formula").Clear();
            api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J+K+L+M)/13").Exe();
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Ana.Length)
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();    //ASCII 65 = "A"
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("Calc_" + CalcTagName_Ana[n - 1]).Exe();
                }
                else
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("").Exe();
                }
                if (n == 20)
                    api.ByName("T").Enter("").Submit().Exe();
            }
        }

        private void CreateCal_Average_OPCUA_Tag()
        {
            //// AvgOPCUA020~AvgOPCUA250 ////
            int iTagCount = 0;
            api.ByName("ParaName").SelectVal("CalcAna").Exe();
            string[] CalcTagName_Ana = { "AvgOPCUA020", "AvgOPCUA040", "AvgOPCUA060", "AvgOPCUA080", "AvgOPCUA100", "AvgOPCUA120", "AvgOPCUA140",
                                         "AvgOPCUA160", "AvgOPCUA180", "AvgOPCUA200", "AvgOPCUA220", "AvgOPCUA240", "AvgOPCUA250"};
            Thread.Sleep(500);
            SetupBasicAnalogTagConfig();
            for (int i = 0; i < CalcTagName_Ana.Length; i++)
            {
                try
                {
                    api.ByName("TagName").Clear();
                    api.ByName("TagName").Enter("Calc_" + CalcTagName_Ana[i]).Exe();

                    api.ByName("Formula").Clear();
                    if (CalcTagName_Ana[i] == "AvgOPCUA250")
                        api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J)/10").Exe();
                    else
                        api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J+K+L+M+N+O+P+Q+R+S+T)/20").Exe();

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Ana[i] == "AvgOPCUA250")
                        {
                            if (j <= 10)
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("OPCUA_" + (iTagCount + j).ToString("0000")).Exe();
                            }
                            else
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("").Exe();
                            }
                        }
                        else
                        {
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("OPCUA_" + (iTagCount + j).ToString("0000")).Exe();
                        }

                        if (j == 20)
                            api.ByName("T").Enter("").Submit().Exe();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create CalcTags error: " + ex.ToString());
                    i--;
                }

                iTagCount += 20;
            }

            //// AvgOPCUAAll ////
            api.ByName("TagName").Clear();
            api.ByName("TagName").Enter("Calc_AvgOPCUAAll").Exe();
            api.ByName("Formula").Clear();
            api.ByName("Formula").Enter("(A+B+C+D+E+F+G+H+I+J+K+L+M)/13").Exe();
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Ana.Length)
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();    //ASCII 65 = "A"
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("Calc_" + CalcTagName_Ana[n - 1]).Exe();
                }
                else
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("").Exe();
                }
                if (n == 20)
                    api.ByName("T").Enter("").Submit().Exe();
            }
        }

        private void CreateCal_AND_DI_Tag()
        {
            //// ANDDI020~ANDDI250 ////
            int iTagCount = 0;
            api.ByName("ParaName").SelectVal("CalcDis").Exe();
            string[] CalcTagName_Dis = { "ANDDI020", "ANDDI040", "ANDDI060", "ANDDI080", "ANDDI100", "ANDDI120", "ANDDI140",
                                         "ANDDI160", "ANDDI180", "ANDDI200", "ANDDI220", "ANDDI240", "ANDDI250"};
            Thread.Sleep(500);
            SetupBasicDiscreteTagConfig();

            for (int i = 0; i < CalcTagName_Dis.Length; i++)
            {
                try
                {
                    api.ByName("TagName").Clear();
                    api.ByName("TagName").Enter("Calc_" + CalcTagName_Dis[i]).Exe();

                    api.ByName("Formula").Clear();
                    if (CalcTagName_Dis[i] == "ANDDI250")
                        api.ByName("Formula").Enter("A&&B&&C&&D&&E&&F&&G&&H&&I&&J").Exe();
                    else
                        api.ByName("Formula").Enter("A&&B&&C&&D&&E&&F&&G&&H&&I&&J&&K&&L&&M&&N&&O&&P&&Q&&R&&S&&T").Exe();

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Dis[i] == "ANDDI250")
                        {
                            if (j <= 10)
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("AT_DI0" + (iTagCount + j).ToString("000")).Exe();
                            }
                            else
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("").Exe();
                            }
                        }
                        else
                        {
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("AT_DI0" + (iTagCount + j).ToString("000")).Exe();
                        }

                        if (j == 20)
                            api.ByName("T").Enter("").Submit().Exe();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create CalcTags error: " + ex.ToString());
                    i--;
                }

                iTagCount += 20;
            }

            //// ANDDIAll ////
            api.ByName("TagName").Clear();
            api.ByName("TagName").Enter("Calc_ANDDIAll").Exe();
            api.ByName("Formula").Clear();
            api.ByName("Formula").Enter("A&&B&&C&&D&&E&&F&&G&&H&&I&&J&&K&&L&&M").Exe();
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Dis.Length)
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();    //ASCII 65 = "A"
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("Calc_" + CalcTagName_Dis[n - 1]).Exe();
                }
                else
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("").Exe();
                }
                if (n == 20)
                    api.ByName("T").Enter("").Submit().Exe();
            }
        }

        private void CreateCal_AND_DO_Tag()
        {
            //// ANDDO020~ANDDO250 ////
            int iTagCount = 0;
            api.ByName("ParaName").SelectVal("CalcDis").Exe();
            string[] CalcTagName_Dis = { "ANDDO020", "ANDDO040", "ANDDO060", "ANDDO080", "ANDDO100", "ANDDO120", "ANDDO140",
                                         "ANDDO160", "ANDDO180", "ANDDO200", "ANDDO220", "ANDDO240", "ANDDO250"};
            Thread.Sleep(500);
            SetupBasicDiscreteTagConfig();

            for (int i = 0; i < CalcTagName_Dis.Length; i++)
            {
                try
                {
                    api.ByName("TagName").Clear();
                    api.ByName("TagName").Enter("Calc_" + CalcTagName_Dis[i]).Exe();

                    api.ByName("Formula").Clear();
                    if (CalcTagName_Dis[i] == "ANDDO250")
                        api.ByName("Formula").Enter("A&&B&&C&&D&&E&&F&&G&&H&&I&&J").Exe();
                    else
                        api.ByName("Formula").Enter("A&&B&&C&&D&&E&&F&&G&&H&&I&&J&&K&&L&&M&&N&&O&&P&&Q&&R&&S&&T").Exe();

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Dis[i] == "ANDDO250")
                        {
                            if (j <= 10)
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("AT_DO0" + (iTagCount + j).ToString("000")).Exe();
                            }
                            else
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("").Exe();
                            }
                        }
                        else
                        {
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("AT_DO0" + (iTagCount + j).ToString("000")).Exe();
                        }

                        if (j == 20)
                            api.ByName("T").Enter("").Submit().Exe();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create CalcTags error: " + ex.ToString());
                    i--;
                }

                iTagCount += 20;
            }

            //// ANDDOAll ////
            api.ByName("TagName").Clear();
            api.ByName("TagName").Enter("Calc_ANDDOAll").Exe();
            api.ByName("Formula").Clear();
            api.ByName("Formula").Enter("A&&B&&C&&D&&E&&F&&G&&H&&I&&J&&K&&L&&M").Exe();
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Dis.Length)
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();    //ASCII 65 = "A"
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("Calc_" + CalcTagName_Dis[n - 1]).Exe();
                }
                else
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("").Exe();
                }
                if (n == 20)
                    api.ByName("T").Enter("").Submit().Exe();
            }
        }

        private void CreateCal_AND_ConDis_Tag()
        {
            //// ANDConDis020~ANDConDis250 ////
            int iTagCount = 0;
            api.ByName("ParaName").SelectVal("CalcDis").Exe();
            string[] CalcTagName_Dis = { "ANDConDis020", "ANDConDis040", "ANDConDis060", "ANDConDis080", "ANDConDis100", "ANDConDis120", "ANDConDis140",
                                         "ANDConDis160", "ANDConDis180", "ANDConDis200", "ANDConDis220", "ANDConDis240", "ANDConDis250"};
            Thread.Sleep(500);
            SetupBasicDiscreteTagConfig();

            for (int i = 0; i < CalcTagName_Dis.Length; i++)
            {
                try
                {
                    api.ByName("TagName").Clear();
                    api.ByName("TagName").Enter("Calc_" + CalcTagName_Dis[i]).Exe();

                    api.ByName("Formula").Clear();
                    if (CalcTagName_Dis[i] == "ANDConDis250")
                        api.ByName("Formula").Enter("A&&B&&C&&D&&E&&F&&G&&H&&I&&J").Exe();
                    else
                        api.ByName("Formula").Enter("A&&B&&C&&D&&E&&F&&G&&H&&I&&J&&K&&L&&M&&N&&O&&P&&Q&&R&&S&&T").Exe();

                    for (int j = 1; j <= 20; j++)
                    {
                        if (CalcTagName_Dis[i] == "ANDConDis250")
                        {
                            if (j <= 10)
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("ConDis_" + (iTagCount + j).ToString("0000")).Exe();
                            }
                            else
                            {
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();
                                api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("").Exe();
                            }
                        }
                        else
                        {
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Clear();    //ASCII 65 = "A"
                            api.ByName(System.Convert.ToString(Convert.ToChar(j + 64))).Enter("ConDis_" + (iTagCount + j).ToString("0000")).Exe();
                        }

                        if (j == 20)
                            api.ByName("T").Enter("").Submit().Exe();
                    }
                }
                catch (Exception ex)
                {
                    EventLog.AddLog("Create CalcTags error: " + ex.ToString());
                    i--;
                }

                iTagCount += 20;
            }

            //// ANDConDisAll ////
            api.ByName("TagName").Clear();
            api.ByName("TagName").Enter("Calc_ANDConDisAll").Exe();
            api.ByName("Formula").Clear();
            api.ByName("Formula").Enter("A&&B&&C&&D&&E&&F&&G&&H&&I&&J&&K&&L&&M").Exe();
            for (int n = 1; n <= 20; n++)
            {
                if (n <= CalcTagName_Dis.Length)
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();    //ASCII 65 = "A"
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("Calc_" + CalcTagName_Dis[n - 1]).Exe();
                }
                else
                {
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Clear();
                    api.ByName(System.Convert.ToString(Convert.ToChar(n + 64))).Enter("").Exe();
                }
                if (n == 20)
                    api.ByName("T").Enter("").Submit().Exe();
            }
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

        private void SetupBasicDiscreteTagConfig()
        {
            api.ByName("Datalog").Click();
            api.ByName("DataLogDB").Clear();
            api.ByName("DataLogDB").Enter("0").Exe();
            api.ByName("ChangeLog").Click();
            api.ByName("ReservedInt1").Click();
        }

        private void Start_Click(object sender, EventArgs e)
        {
            long lErrorCode = 0;
            EventLog.AddLog("===Create calc tags start===");
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, WebAccessIP.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===Create calc tags end===");
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
