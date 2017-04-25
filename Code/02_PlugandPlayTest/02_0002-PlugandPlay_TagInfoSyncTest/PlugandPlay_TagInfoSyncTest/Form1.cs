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
using System.Runtime.InteropServices;
using System.Diagnostics;
using ThirdPartyToolControl;
using iATester;
using WAWebServiceInterface;
using CommonFunction;

namespace PlugandPlay_TagInfoSyncTest
{
    public partial class Form1 : Form, iATester.iCom
    {
        IAdvSeleniumAPI api;
        cThirdPartyToolControl tpc = new cThirdPartyToolControl();
        cWACommonFunction wacf = new cWACommonFunction();
        cEventLog EventLog = new cEventLog();

        WAWebServiceInterface.WAWebService waWebSvcG = null;
        WAWebServiceInterface.WAWebService waWebSvcC = null;

        WAGetValResponseObj G_AI_Value, G_AO_Value, G_DI_Value, G_DO_Value, G_OPCDA_Value, G_OPCUA_Value, G_Acc_Value, G_ConAna_Value, G_ConDis_Value, G_ConTxt_Value, G_Calc_Value, G_System_Value = null;
        WAGetValResponseObj C_AI_Value, C_AO_Value, C_DI_Value, C_DO_Value, C_OPCDA_Value, C_OPCUA_Value, C_Acc_Value, C_ConAna_Value, C_ConDis_Value, C_ConTxt_Value, C_Calc_Value, C_System_Value = null;

        private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
        private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
        internal const int Max_Rows_Val = 65535;
        string baseUrl, baseUrl2;
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
            long lErrorCode = 0;
            EventLog.AddLog("===PlugandPlay_TagInfoSyncTest start (by iATester)===");
            if (System.IO.File.Exists(sIniFilePath))    // 再load一次
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, ProjectName2.Text, WebAccessIP.Text, WebAccessIP2.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===PlugandPlay_TagInfoSyncTest end (by iATester)===");

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

        long Form1_Load(string sProjectName, string sProjectName2, string sWebAccessIP, string sWebAccessIP2, string sTestLogFolder, string sBrowser)
        {
            baseUrl = "http://" + sWebAccessIP;
            baseUrl2 = "http://" + sWebAccessIP2;

            // Step1: Start the kernel of ground PC
            GroundPCStartKernel(sBrowser, sProjectName, sWebAccessIP, sTestLogFolder);

            // Step2: Start the kernel of cloud PC 
            CloudPCStartKernel(sBrowser, sProjectName2, sWebAccessIP2, sTestLogFolder);

            // Step3: check the cloud and ground tag value
            bool bTagCheckResult;
            bool bGetTagInfoRes = GetTagInfo(sProjectName, sWebAccessIP, sProjectName2, sWebAccessIP2);
            if (bGetTagInfoRes)
                bTagCheckResult = CheckGroundandCloudTag();
            else
                bTagCheckResult = bGetTagInfoRes;

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

            if (bSeleniumResult && bTagCheckResult)
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

        private void GroundPCStartKernel(string sBrowser, string sProjectName, string sWebAccessIP, string sTestLogFolder)
        {
            if (sBrowser == "Internet Explorer")
            {
                EventLog.AddLog("<GroundPC> Browser= Internet Explorer");
                api = new AdvSeleniumAPI("IE", "");
                System.Threading.Thread.Sleep(1000);
            }
            else if (sBrowser == "Mozilla FireFox")
            {
                EventLog.AddLog("<GroundPC> Browser= Mozilla FireFox");
                api = new AdvSeleniumAPI("FireFox", "");
                System.Threading.Thread.Sleep(1000);
            }

            api.LinkWebUI(baseUrl + "/broadWeb/bwconfig.asp?username=admin");
            api.ById("userField").Enter("").Submit().Exe();
            PrintStep(api, "<GroundPC> Login WebAccess");

            // Configure project by project name
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
            PrintStep(api, "<GroundPC> Configure project");

            EventLog.AddLog("<GroundPC> Start Kernel");
            wacf.StartKernel(api);

            api.Quit();
            PrintStep(api, "<GroundPC> Quit browser");
        }

        private void CloudPCStartKernel(string sBrowser, string sProjectName, string sWebAccessIP, string sTestLogFolder)
        {
            if (sBrowser == "Internet Explorer")
            {
                EventLog.AddLog("<CloudPC> Browser= Internet Explorer");
                api = new AdvSeleniumAPI("IE", "");
                System.Threading.Thread.Sleep(1000);
            }
            else if (sBrowser == "Mozilla FireFox")
            {
                EventLog.AddLog("<CloudPC> Browser= Mozilla FireFox");
                api = new AdvSeleniumAPI("FireFox", "");
                System.Threading.Thread.Sleep(1000);
            }

            api.LinkWebUI(baseUrl2 + "/broadWeb/bwconfig.asp?username=admin");
            api.ById("userField").Enter("").Submit().Exe();
            PrintStep(api, "<CloudPC> Login WebAccess");

            // Configure project by project name
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
            PrintStep(api, "<CloudPC> Configure project");

            EventLog.AddLog("<CloudPC> Start Kernel");
            wacf.StartKernel(api);

            api.Quit();
            PrintStep(api, "<CloudPC> Quit browser");
        }

        private void GetCloudTagInfo()
        {
            EventLog.AddLog("Get cloud tag info");
            /*
            string user = "admin";
            string pwd = "";
            if (null == waWebSvcC) waWebSvcC = new WAWebService();
            EventLog.AddLog("Send GetProjectList Request");

            if (waWebSvcC.Init(sWebAccessIP, "", user, pwd))
            {
                EventLog.AddLog("waWebSvc.Init() Success");
            }
            else
                EventLog.AddLog("waWebSvc.Init() Fail");
            
            waWebSvcC.SetProject(sProjectName);      // set Project name first
            */

            // Get AI value
            string[] sAITagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sAITagList[i - 1] = "AT_AI" + i.ToString("0000");
            C_AI_Value = waWebSvcC.GetValueText(sAITagList, false);
            //for (int i = 0; i < C_AI_Value.Result.Total; i++)
            //{
            //    string sName = C_AI_Value.Values[i].Name;
            //    string sValues = C_AI_Value.Values[i].Value;
            //    int iQuality = C_AI_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get AO value
            string[] sAOTagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sAOTagList[i - 1] = "AT_AO" + i.ToString("0000");
            C_AO_Value = waWebSvcC.GetValueText(sAOTagList, false);
            //for (int i = 0; i < C_AO_Value.Result.Total; i++)
            //{
            //    string sName = C_AO_Value.Values[i].Name;
            //    string sValues = C_AO_Value.Values[i].Value;
            //    int iQuality = C_AO_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}
            
            // Get DI value
            string[] sDITagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sDITagList[i - 1] = "AT_DI" + i.ToString("0000");
            C_DI_Value = waWebSvcC.GetValueText(sDITagList, false);
            //for (int i = 0; i < C_DI_Value.Result.Total; i++)
            //{
            //    string sName = C_DI_Value.Values[i].Name;
            //    string sValues = C_DI_Value.Values[i].Value;
            //    int iQuality = C_DI_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get DO value
            string[] sDOTagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sDOTagList[i - 1] = "AT_DO" + i.ToString("0000");
            C_DO_Value = waWebSvcC.GetValueText(sDOTagList, false);
            //for (int i = 0; i < C_DO_Value.Result.Total; i++)
            //{
            //    string sName = C_DO_Value.Values[i].Name;
            //    string sValues = C_DO_Value.Values[i].Value;
            //    int iQuality = C_DO_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get OPCDA value
            string[] sOPCDATagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sOPCDATagList[i - 1] = "OPCDA_" + i.ToString("0000");
            C_OPCDA_Value = waWebSvcC.GetValueText(sOPCDATagList, false);
            //for (int i = 0; i < C_OPCDA_Value.Result.Total; i++)
            //{
            //    string sName = C_OPCDA_Value.Values[i].Name;
            //    string sValues = C_OPCDA_Value.Values[i].Value;
            //    int iQuality = C_OPCDA_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get OPCUA value
            string[] sOPCUATagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sOPCUATagList[i - 1] = "OPCUA_" + i.ToString("0000");
            C_OPCUA_Value = waWebSvcC.GetValueText(sOPCUATagList, false);
            //for (int i = 0; i < C_OPCUA_Value.Result.Total; i++)
            //{
            //    string sName = C_OPCUA_Value.Values[i].Name;
            //    string sValues = C_OPCUA_Value.Values[i].Value;
            //    int iQuality = C_OPCUA_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get Acc value
            string[] sAccTagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sAccTagList[i - 1] = "Acc_" + i.ToString("0000");
            C_Acc_Value = waWebSvcC.GetValueText(sAccTagList, false);
            //for (int i = 0; i < C_Acc_Value.Result.Total; i++)
            //{
            //    string sName = C_Acc_Value.Values[i].Name;
            //    string sValues = C_Acc_Value.Values[i].Value;
            //    int iQuality = C_Acc_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get ConAna value
            string[] sConAnaTagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sConAnaTagList[i - 1] = "ConAna_" + i.ToString("0000");
            C_ConAna_Value = waWebSvcC.GetValueText(sConAnaTagList, false);
            //for (int i = 0; i < C_ConAna_Value.Result.Total; i++)
            //{
            //    string sName = C_ConAna_Value.Values[i].Name;
            //    string sValues = C_ConAna_Value.Values[i].Value;
            //    int iQuality = C_ConAna_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get ConDis value
            string[] sConDisTagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sConDisTagList[i - 1] = "ConDis_" + i.ToString("0000");
            C_ConDis_Value = waWebSvcC.GetValueText(sConDisTagList, false);
            //for (int i = 0; i < C_ConDis_Value.Result.Total; i++)
            //{
            //    string sName = C_ConDis_Value.Values[i].Name;
            //    string sValues = C_ConDis_Value.Values[i].Value;
            //    int iQuality = C_ConDis_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get ConTxt value
            string[] sConTxtTagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sConTxtTagList[i - 1] = "ConTxt_" + i.ToString("0000");
            C_ConTxt_Value = waWebSvcC.GetValueText(sConTxtTagList, true);   // set true for Text
            //for (int i = 0; i < C_ConTxt_Value.Result.Total; i++)
            //{
            //    string sName = C_ConTxt_Value.Values[i].Name;
            //    string sValues = C_ConTxt_Value.Values[i].Value;
            //    int iQuality = C_ConTxt_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get Calculate value
            //string[] sCalcTagList = { "Calc_ModBusAI", "Calc_ModBusAO", "Calc_ModBusDI", "Calc_ModBusDO",
            //"Calc_OPCDA", "Calc_OPCUA", "Calc_Acc", "Calc_ConAna", "Calc_ConDis",  "Calc_System"};
            string[] sCalcTagList = { "Calc_ModBusAI", "Calc_ModBusAO", "Calc_OPCDA", "Calc_OPCUA", "Calc_Acc", "Calc_ConAna",  "Calc_System"};
            C_Calc_Value = waWebSvcC.GetValueText(sCalcTagList, false);
            //for (int i = 0; i < C_Calc_Value.Result.Total; i++)
            //{
            //    string sName = C_Calc_Value.Values[i].Name;
            //    string sValues = C_Calc_Value.Values[i].Value;
            //    int iQuality = C_Calc_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get System value
            string[] sSystemTagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sSystemTagList[i - 1] = "SystemSec_" + i.ToString("0000");
            C_System_Value = waWebSvcC.GetValueText(sSystemTagList, false);
            //for (int i = 0; i < C_System_Value.Result.Total; i++)
            //{
            //    string sName = C_System_Value.Values[i].Name;
            //    string sValues = C_System_Value.Values[i].Value;
            //    int iQuality = C_System_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}
            EventLog.AddLog("Get cloud tag info END");
        }

        private void GetGroundTagInfo()
        {
            EventLog.AddLog("Get ground tag info");
            /*
            string user = "admin";
            string pwd = "";
            if (null == waWebSvcG) waWebSvcG = new WAWebService();
            EventLog.AddLog("Send GetProjectList Request");

            if (waWebSvcG.Init(sWebAccessIP, "", user, pwd))
            {
                EventLog.AddLog("waWebSvc.Init() Success");
            }
            else
                EventLog.AddLog("waWebSvc.Init() Fail");

            waWebSvcG.SetProject(sProjectName);      // set Project name first
            */
            // Get AI value
            string[] sAITagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sAITagList[i - 1] = "AT_AI" + i.ToString("0000");
            G_AI_Value = waWebSvcG.GetValueText(sAITagList, false);
            //for (int i = 0; i < G_AI_Value.Result.Total; i++)
            //{
            //    string sName = G_AI_Value.Values[i].Name;
            //    string sValues = G_AI_Value.Values[i].Value;
            //    int iQuality = G_AI_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get AO value
            string[] sAOTagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sAOTagList[i - 1] = "AT_AO" + i.ToString("0000");
            G_AO_Value = waWebSvcG.GetValueText(sAOTagList, false);
            //for (int i = 0; i < G_AO_Value.Result.Total; i++)
            //{
            //    string sName = G_AO_Value.Values[i].Name;
            //    string sValues = G_AO_Value.Values[i].Value;
            //    int iQuality = G_AO_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get DI value
            string[] sDITagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sDITagList[i - 1] = "AT_DI" + i.ToString("0000");
            G_DI_Value = waWebSvcG.GetValueText(sDITagList, false);
            //for (int i = 0; i < G_DI_Value.Result.Total; i++)
            //{
            //    string sName = G_DI_Value.Values[i].Name;
            //    string sValues = G_DI_Value.Values[i].Value;
            //    int iQuality = G_DI_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get DO value
            string[] sDOTagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sDOTagList[i - 1] = "AT_DO" + i.ToString("0000");
            G_DO_Value = waWebSvcG.GetValueText(sDOTagList, false);
            //for (int i = 0; i < G_DO_Value.Result.Total; i++)
            //{
            //    string sName = G_DO_Value.Values[i].Name;
            //    string sValues = G_DO_Value.Values[i].Value;
            //    int iQuality = G_DO_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get OPCDA value
            string[] sOPCDATagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sOPCDATagList[i - 1] = "OPCDA_" + i.ToString("0000");
            G_OPCDA_Value = waWebSvcG.GetValueText(sOPCDATagList, false);
            //for (int i = 0; i < G_OPCDA_Value.Result.Total; i++)
            //{
            //    string sName = G_OPCDA_Value.Values[i].Name;
            //    string sValues = G_OPCDA_Value.Values[i].Value;
            //    int iQuality = G_OPCDA_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get OPCUA value
            string[] sOPCUATagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sOPCUATagList[i - 1] = "OPCUA_" + i.ToString("0000");
            G_OPCUA_Value = waWebSvcG.GetValueText(sOPCUATagList, false);
            //for (int i = 0; i < G_OPCUA_Value.Result.Total; i++)
            //{
            //    string sName = G_OPCUA_Value.Values[i].Name;
            //    string sValues = G_OPCUA_Value.Values[i].Value;
            //    int iQuality = G_OPCUA_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get Acc value
            string[] sAccTagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sAccTagList[i - 1] = "Acc_" + i.ToString("0000");
            G_Acc_Value = waWebSvcG.GetValueText(sAccTagList, false);
            //for (int i = 0; i < G_Acc_Value.Result.Total; i++)
            //{
            //    string sName = G_Acc_Value.Values[i].Name;
            //    string sValues = G_Acc_Value.Values[i].Value;
            //    int iQuality = G_Acc_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get ConAna value
            string[] sConAnaTagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sConAnaTagList[i - 1] = "ConAna_" + i.ToString("0000");
            G_ConAna_Value = waWebSvcG.GetValueText(sConAnaTagList, false);
            //for (int i = 0; i < G_ConAna_Value.Result.Total; i++)
            //{
            //    string sName = G_ConAna_Value.Values[i].Name;
            //    string sValues = G_ConAna_Value.Values[i].Value;
            //    int iQuality = G_ConAna_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get ConDis value
            string[] sConDisTagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sConDisTagList[i - 1] = "ConDis_" + i.ToString("0000");
            G_ConDis_Value = waWebSvcG.GetValueText(sConDisTagList, false);
            //for (int i = 0; i < G_ConDis_Value.Result.Total; i++)
            //{
            //    string sName = G_ConDis_Value.Values[i].Name;
            //    string sValues = G_ConDis_Value.Values[i].Value;
            //    int iQuality = G_ConDis_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get ConTxt value
            string[] sConTxtTagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sConTxtTagList[i - 1] = "ConTxt_" + i.ToString("0000");
            G_ConTxt_Value = waWebSvcG.GetValueText(sConTxtTagList, true);   // set true for Text
            //for (int i = 0; i < G_ConTxt_Value.Result.Total; i++)
            //{
            //    string sName = G_ConTxt_Value.Values[i].Name;
            //    string sValues = G_ConTxt_Value.Values[i].Value;
            //    int iQuality = G_ConTxt_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get Calculate value
            //string[] sCalcTagList = { "Calc_ModBusAI", "Calc_ModBusAO", "Calc_ModBusDI", "Calc_ModBusDO",
            //"Calc_OPCDA", "Calc_OPCUA", "Calc_Acc", "Calc_ConAna", "Calc_ConDis",  "Calc_System"};
            string[] sCalcTagList = { "Calc_ModBusAI", "Calc_ModBusAO", "Calc_OPCDA", "Calc_OPCUA", "Calc_Acc", "Calc_ConAna",  "Calc_System"};
            G_Calc_Value = waWebSvcG.GetValueText(sCalcTagList, false);
            //for (int i = 0; i < G_Calc_Value.Result.Total; i++)
            //{
            //    string sName = G_Calc_Value.Values[i].Name;
            //    string sValues = G_Calc_Value.Values[i].Value;
            //    int iQuality = G_Calc_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}

            // Get System value
            string[] sSystemTagList = new string[250];
            for (int i = 1; i <= 250; i++)
                sSystemTagList[i - 1] = "SystemSec_" + i.ToString("0000");
            G_System_Value = waWebSvcG.GetValueText(sSystemTagList, false);
            //for (int i = 0; i < G_System_Value.Result.Total; i++)
            //{
            //    string sName = G_System_Value.Values[i].Name;
            //    string sValues = G_System_Value.Values[i].Value;
            //    int iQuality = G_System_Value.Values[i].Quality;
            //    EventLog.AddLog("TagName: " + sName + " TagValue: " + sValues + string.Format(" Quality: {0}", iQuality));
            //}
            EventLog.AddLog("Get ground tag info END");
        }

        private bool GetTagInfo(string sProjectName, string sWebAccessIP, string sProjectName2, string sWebAccessIP2)
        {
            string user = "admin";
            string pwd = "";
            bool bGetTagInfoResult;
            if (null == waWebSvcC) waWebSvcC = new WAWebService();
            if (null == waWebSvcG) waWebSvcG = new WAWebService();
            
            bool bInitGround = waWebSvcG.Init(sWebAccessIP, sProjectName, user, pwd);
            bool bInitCloud = waWebSvcC.Init(sWebAccessIP2, sProjectName2, user, pwd);

            if (!(bInitGround && bInitGround))
            {
                if(!bInitGround)
                {
                    EventLog.AddLog("Ground waWebSvc.Init() Fail!!");
                    EventLog.AddLog("Error message: " + waWebSvcG.GetErrMsg());
                }
                if(!bInitCloud)
                {
                    EventLog.AddLog("Cloud waWebSvc.Init() Fail!!");
                    EventLog.AddLog("Cloud error message: " + waWebSvcC.GetErrMsg());
                }
                bGetTagInfoResult = false;
            }
            else
            {
                EventLog.AddLog("Ground and Cloud waWebSvc.Init() Success");

                // Start get ground and cloud tag value
                EventLog.AddLog("Get ground and cloud tag info start!!");

                // Get AI value
                EventLog.AddLog("Get AI value..");
                string[] sAITagList = new string[250];
                for (int i = 1; i <= 250; i++)
                    sAITagList[i - 1] = "AT_AI" + i.ToString("0000");
                EventLog.AddLog("Get ground 250 AI tag start");
                G_AI_Value = waWebSvcG.GetValueText(sAITagList, false);
                EventLog.AddLog("Get ground 250 AI tag end");
                EventLog.AddLog("Get cloud 250 AI tag start");
                C_AI_Value = waWebSvcC.GetValueText(sAITagList, false);
                EventLog.AddLog("Get cloud 250 AI tag end");

                // Get AO value
                EventLog.AddLog("Get AO value..");
                string[] sAOTagList = new string[250];
                for (int i = 1; i <= 250; i++)
                    sAOTagList[i - 1] = "AT_AO" + i.ToString("0000");
                G_AO_Value = waWebSvcG.GetValueText(sAOTagList, false);
                C_AO_Value = waWebSvcC.GetValueText(sAOTagList, false);

                // Get DI value
                EventLog.AddLog("Get DI value..");
                string[] sDITagList = new string[250];
                for (int i = 1; i <= 250; i++)
                    sDITagList[i - 1] = "AT_DI" + i.ToString("0000");
                G_DI_Value = waWebSvcG.GetValueText(sDITagList, false);
                C_DI_Value = waWebSvcC.GetValueText(sDITagList, false);

                // Get DO value
                EventLog.AddLog("Get DO value..");
                string[] sDOTagList = new string[250];
                for (int i = 1; i <= 250; i++)
                    sDOTagList[i - 1] = "AT_DO" + i.ToString("0000");
                G_DO_Value = waWebSvcG.GetValueText(sDOTagList, false);
                C_DO_Value = waWebSvcC.GetValueText(sDOTagList, false);

                // Get OPCDA value
                EventLog.AddLog("Get OPCDA value..");
                string[] sOPCDATagList = new string[250];
                for (int i = 1; i <= 250; i++)
                    sOPCDATagList[i - 1] = "OPCDA_" + i.ToString("0000");
                G_OPCDA_Value = waWebSvcG.GetValueText(sOPCDATagList, false);
                C_OPCDA_Value = waWebSvcC.GetValueText(sOPCDATagList, false);

                // Get OPCUA value
                EventLog.AddLog("Get OPCUA value..");
                string[] sOPCUATagList = new string[250];
                for (int i = 1; i <= 250; i++)
                    sOPCUATagList[i - 1] = "OPCUA_" + i.ToString("0000");
                G_OPCUA_Value = waWebSvcG.GetValueText(sOPCUATagList, false);
                C_OPCUA_Value = waWebSvcC.GetValueText(sOPCUATagList, false);

                // Get Acc value
                EventLog.AddLog("Get Acc value..");
                string[] sAccTagList = new string[250];
                for (int i = 1; i <= 250; i++)
                    sAccTagList[i - 1] = "Acc_" + i.ToString("0000");
                G_Acc_Value = waWebSvcG.GetValueText(sAccTagList, false);
                C_Acc_Value = waWebSvcC.GetValueText(sAccTagList, false);

                // Get ConAna value
                EventLog.AddLog("Get ConAna value..");
                string[] sConAnaTagList = new string[250];
                for (int i = 1; i <= 250; i++)
                    sConAnaTagList[i - 1] = "ConAna_" + i.ToString("0000");
                G_ConAna_Value = waWebSvcG.GetValueText(sConAnaTagList, false);
                C_ConAna_Value = waWebSvcC.GetValueText(sConAnaTagList, false);

                // Get ConDis value
                EventLog.AddLog("Get ConDis value..");
                string[] sConDisTagList = new string[250];
                for (int i = 1; i <= 250; i++)
                    sConDisTagList[i - 1] = "ConDis_" + i.ToString("0000");
                G_ConDis_Value = waWebSvcG.GetValueText(sConDisTagList, false);
                C_ConDis_Value = waWebSvcC.GetValueText(sConDisTagList, false);

                // Get ConTxt value
                EventLog.AddLog("Get ConTxt value..");
                string[] sConTxtTagList = new string[250];
                for (int i = 1; i <= 250; i++)
                    sConTxtTagList[i - 1] = "ConTxt_" + i.ToString("0000");
                G_ConTxt_Value = waWebSvcG.GetValueText(sConTxtTagList, true);   // set true for Text
                C_ConTxt_Value = waWebSvcC.GetValueText(sConTxtTagList, true);   // set true for Text

                // Get Calculate value
                EventLog.AddLog("Get Calculate value..");
                //string[] sCalcTagList = { "Calc_ModBusAI", "Calc_ModBusAO", "Calc_ModBusDI", "Calc_ModBusDO",
                //                          "Calc_OPCDA", "Calc_OPCUA", "Calc_Acc", "Calc_ConAna", "Calc_ConDis",  "Calc_System"};
                string[] sCalcTagList = { "Calc_ModBusAI", "Calc_ModBusAO", "Calc_OPCDA", "Calc_OPCUA", "Calc_Acc", "Calc_ConAna",  "Calc_System"};
                G_Calc_Value = waWebSvcG.GetValueText(sCalcTagList, false);
                C_Calc_Value = waWebSvcC.GetValueText(sCalcTagList, false);

                // Get System value
                EventLog.AddLog("Get System value..");
                string[] sSystemTagList = new string[250];
                for (int i = 1; i <= 250; i++)
                    sSystemTagList[i - 1] = "SystemSec_" + i.ToString("0000");
                G_System_Value = waWebSvcG.GetValueText(sSystemTagList, false);
                C_System_Value = waWebSvcC.GetValueText(sSystemTagList, false);

                EventLog.AddLog("Get ground and cloud tag info end!!");

                bGetTagInfoResult = true;
            }

            return bGetTagInfoResult;
        }

        private bool CheckGroundandCloudTag()
        {
            EventLog.AddLog("Start check ground and cloud tag difference!!");
            int iAnalogDiffSpec = 2;
            EventLog.AddLog(string.Format("Difference spec of analog tag= {0}",iAnalogDiffSpec));

            // Check AI value
            EventLog.AddLog("Check AI value..");
            bool bAI_Check = true;
            for (int i = 0; i < C_AI_Value.Result.Total; i++)
            {
                double dDiff = Convert.ToDouble(C_AI_Value.Values[i].Value) - Convert.ToDouble(G_AI_Value.Values[i].Value);
                if (Math.Abs(dDiff) > iAnalogDiffSpec)
                {
                    EventLog.AddLog(string.Format("The difference of cloud value({0}={1}) and ground value({2}={3}) is out of spec({4})"
                        , C_AI_Value.Values[i].Name, C_AI_Value.Values[i].Value, G_AI_Value.Values[i].Name, G_AI_Value.Values[i].Value, iAnalogDiffSpec));
                    bAI_Check = false;
                }
            }

            // Check AO value
            EventLog.AddLog("Check AO value..");
            bool bAO_Check = true;
            for (int i = 0; i < C_AO_Value.Result.Total; i++)
            {
                double dDiff = Convert.ToDouble(C_AO_Value.Values[i].Value) - Convert.ToDouble(G_AO_Value.Values[i].Value);
                if (Math.Abs(dDiff) > iAnalogDiffSpec)
                {
                    EventLog.AddLog(string.Format("The difference of cloud value({0}={1}) and ground value({2}={3}) is out of spec({4})"
                        , C_AO_Value.Values[i].Name, C_AO_Value.Values[i].Value, G_AO_Value.Values[i].Name, G_AO_Value.Values[i].Value, iAnalogDiffSpec));
                    bAO_Check = false;
                }
            }

            // Check DI value
            EventLog.AddLog("Check DI value..");
            bool bDI_Check = true;
            for (int i = 0; i < C_DI_Value.Result.Total; i++)
            {
                double dDiff = Convert.ToDouble(C_DI_Value.Values[i].Value) - Convert.ToDouble(G_DI_Value.Values[i].Value);
                if (Math.Abs(dDiff) != 1 && Math.Abs(dDiff) != 0)
                {
                    EventLog.AddLog(string.Format("The difference of cloud value({0}={1}) and ground value({2}={3}) is out of spec(1 or 0)"
                        , C_DI_Value.Values[i].Name, C_DI_Value.Values[i].Value, G_DI_Value.Values[i].Name, G_DI_Value.Values[i].Value));
                    bDI_Check = false;
                }
            }

            // Check DO value
            EventLog.AddLog("Check DO value..");
            bool bDO_Check = true;
            for (int i = 0; i < C_DO_Value.Result.Total; i++)
            {
                double dDiff = Convert.ToDouble(C_DO_Value.Values[i].Value) - Convert.ToDouble(G_DO_Value.Values[i].Value);
                if (Math.Abs(dDiff) != 1 && Math.Abs(dDiff) != 0)
                {
                    EventLog.AddLog(string.Format("The difference of cloud value({0}={1}) and ground value({2}={3}) is out of spec(1 or 0)"
                        , C_DO_Value.Values[i].Name, C_DO_Value.Values[i].Value, G_DO_Value.Values[i].Name, G_DO_Value.Values[i].Value));
                    bDO_Check = false;
                }
            }

            // Check OPCDA value
            EventLog.AddLog("Check OPCDA value..");
            bool bOPCDA_Check = true;
            for (int i = 0; i < C_OPCDA_Value.Result.Total; i++)
            {
                double dDiff = Convert.ToDouble(C_OPCDA_Value.Values[i].Value) - Convert.ToDouble(G_OPCDA_Value.Values[i].Value);
                if (Math.Abs(dDiff) > iAnalogDiffSpec)
                {
                    EventLog.AddLog(string.Format("The difference of cloud value({0}={1}) and ground value({2}={3}) is out of spec({4})"
                        , C_OPCDA_Value.Values[i].Name, C_OPCDA_Value.Values[i].Value, G_OPCDA_Value.Values[i].Name, G_OPCDA_Value.Values[i].Value, iAnalogDiffSpec));
                    bOPCDA_Check = false;
                }
            }

            // Check OPCUA value
            EventLog.AddLog("Check OPCUA value..");
            bool bOPCUA_Check = true;
            for (int i = 0; i < C_OPCUA_Value.Result.Total; i++)
            {
                double dDiff = Convert.ToDouble(C_OPCUA_Value.Values[i].Value) - Convert.ToDouble(G_OPCUA_Value.Values[i].Value);
                if (Math.Abs(dDiff) > iAnalogDiffSpec)
                {
                    EventLog.AddLog(string.Format("The difference of cloud value({0}={1}) and ground value({2}={3}) is out of spec({4})"
                        , C_OPCUA_Value.Values[i].Name, C_OPCUA_Value.Values[i].Value, G_OPCUA_Value.Values[i].Name, G_OPCUA_Value.Values[i].Value, iAnalogDiffSpec));
                    bOPCUA_Check = false;
                }
            }

            // Check Acc value
            EventLog.AddLog("Check Acc value..");
            bool bAcc_Check = true;
            for (int i = 0; i < C_Acc_Value.Result.Total; i++)
            {
                double dDiff = Convert.ToDouble(C_Acc_Value.Values[i].Value) - Convert.ToDouble(G_Acc_Value.Values[i].Value);
                if (Math.Abs(dDiff) > iAnalogDiffSpec)
                {
                    EventLog.AddLog(string.Format("The difference of cloud value({0}={1}) and ground value({2}={3}) is out of spec({4})"
                        , C_Acc_Value.Values[i].Name, C_Acc_Value.Values[i].Value, G_Acc_Value.Values[i].Name, G_Acc_Value.Values[i].Value, iAnalogDiffSpec));
                    bAcc_Check = false;
                }
            }

            // Check ConAna value
            EventLog.AddLog("Check ConAna value..");
            bool bConAna_Check = true;
            for (int i = 0; i < C_ConAna_Value.Result.Total; i++)
            {
                double dDiff = Convert.ToDouble(C_ConAna_Value.Values[i].Value) - Convert.ToDouble(G_ConAna_Value.Values[i].Value);
                if (Math.Abs(dDiff) > iAnalogDiffSpec)
                {
                    EventLog.AddLog(string.Format("The difference of cloud value({0}={1}) and ground value({2}={3}) is out of spec({4})"
                        , C_ConAna_Value.Values[i].Name, C_ConAna_Value.Values[i].Value, G_ConAna_Value.Values[i].Name, G_ConAna_Value.Values[i].Value, iAnalogDiffSpec));
                    bConAna_Check = false;
                }
            }

            // Check ConDis value
            EventLog.AddLog("Check ConDis value..");
            bool bConDis_Check = true;
            for (int i = 0; i < C_ConDis_Value.Result.Total; i++)
            {
                double dDiff = Convert.ToDouble(C_ConDis_Value.Values[i].Value) - Convert.ToDouble(G_ConDis_Value.Values[i].Value);
                if (Math.Abs(dDiff) != 1 && Math.Abs(dDiff) != 0)
                {
                    EventLog.AddLog(string.Format("The difference of cloud value({0}={1}) and ground value({2}={3}) is out of spec(1 or 0)"
                        , C_ConDis_Value.Values[i].Name, C_ConDis_Value.Values[i].Value, G_ConDis_Value.Values[i].Name, G_ConDis_Value.Values[i].Value));
                    bConDis_Check = false;
                }
            }

            // Check ConTxt value
            EventLog.AddLog("Check ConTxt value..");
            bool bConTxt_Check = true;
            for (int i = 0; i < C_ConTxt_Value.Result.Total; i++)
            {
                if (C_ConTxt_Value.Values[i].Value != G_ConTxt_Value.Values[i].Value)
                {
                    EventLog.AddLog(string.Format("The difference of cloud value({0}={1}) and ground value({2}={3}) is not equal"
                        , C_ConTxt_Value.Values[i].Name, C_ConTxt_Value.Values[i].Value, G_ConTxt_Value.Values[i].Name, G_ConTxt_Value.Values[i].Value));
                    bConTxt_Check = false;
                }
            }

            // Check Calculate value
            EventLog.AddLog("Check Calculate value..");
            bool bCalc_Check = true;
            for (int i = 0; i < C_Calc_Value.Result.Total; i++)
            {
                double dDiff = Convert.ToDouble(C_Calc_Value.Values[i].Value) - Convert.ToDouble(G_Calc_Value.Values[i].Value);
                if (Math.Abs(dDiff) > iAnalogDiffSpec)
                {
                    EventLog.AddLog(string.Format("The difference of cloud value({0}={1}) and ground value({2}={3}) is out of spec({4})"
                        , C_Calc_Value.Values[i].Name, C_Calc_Value.Values[i].Value, G_Calc_Value.Values[i].Name, G_Calc_Value.Values[i].Value, iAnalogDiffSpec));
                    bCalc_Check = false;
                }
            }

            // Check System value
            EventLog.AddLog("Check System value..");
            bool bSystem_Check = true;
            for (int i = 0; i < C_System_Value.Result.Total; i++)
            {
                double dDiff = Convert.ToDouble(C_System_Value.Values[i].Value) - Convert.ToDouble(G_System_Value.Values[i].Value);
                if ((Math.Abs(dDiff) > iAnalogDiffSpec) && 
                    (Math.Abs(dDiff) != 60-iAnalogDiffSpec) && 
                    (Math.Abs(dDiff) != 60-iAnalogDiffSpec+1)) // 秒數會從59跳0 從0開始累加, 避免有58 vs 0或59 vs 0的情況 而誤判
                {
                    EventLog.AddLog(string.Format("The difference of cloud value({0}={1}) and ground value({2}={3}) is out of spec({4})"
                        , C_System_Value.Values[i].Name, C_System_Value.Values[i].Value, G_System_Value.Values[i].Name, G_System_Value.Values[i].Value, iAnalogDiffSpec));
                    bSystem_Check = false;
                }
            }

            bool bTotalResult = bAI_Check && bAO_Check && bDI_Check && bDO_Check && bOPCDA_Check && bOPCUA_Check 
                && bAcc_Check && bConAna_Check && bConDis_Check && bConTxt_Check && bCalc_Check && bSystem_Check;

            if (!bTotalResult)
                EventLog.AddLog("Check ground and cloud tag difference FAILL!!");
            else
                EventLog.AddLog("Check ground and cloud tag difference PASS!!");

            return bTotalResult;
        }

        private void SendCharToHandle(int iHandle, int iDelay, string sText)
        {
            var chars = sText.ToCharArray();
            for (int ctr = 0; ctr < chars.Length; ctr++)
            {
                tpc.F_PostMessage(iHandle, tpc.V_WM_CHAR, chars[ctr], 0);
                Thread.Sleep(iDelay);
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

        private void ReturnSCADAPage(IAdvSeleniumAPI api)
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMainRight.asp') and contains(@href, 'name=TestSCADA')]").Click();

        }

        private void PrintStep(IAdvSeleniumAPI api, string sTestItem)
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

        private void Start_Click(object sender, EventArgs e)
        {
            long lErrorCode = 0;
            EventLog.AddLog("===PlugandPlay_TagInfoSyncTest start===");
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address(Ground PC)= " + WebAccessIP.Text);
            EventLog.AddLog("WebAccess IP address(Cloud PC)= " + WebAccessIP2.Text);
            lErrorCode = Form1_Load(ProjectName.Text, ProjectName2.Text, WebAccessIP.Text, WebAccessIP2.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===PlugandPlay_TagInfoSyncTest end===");
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

            ProjectName2.Text = sDefaultProjectName2.ToString();
            WebAccessIP2.Text = sDefaultIP2.ToString();
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
                if (ProjectName2.Text != sDefaultProjectName2.ToString())
                {
                    tpc.F_WritePrivateProfileString("ProjectName", "Cloud PC or Backup PC", ProjectName2.Text, sIniFilePath);
                    EventLog.AddLog("New ProjectName update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultProjectName2.ToString());
                    EventLog.AddLog("New ini:" + ProjectName2.Text);
                }
                if (WebAccessIP2.Text != sDefaultIP2.ToString())
                {
                    tpc.F_WritePrivateProfileString("IP", "Cloud PC or Backup PC", WebAccessIP2.Text, sIniFilePath);
                    EventLog.AddLog("New WebAccessIP update to .ini file!!");
                    EventLog.AddLog("Original ini:" + sDefaultIP2.ToString());
                    EventLog.AddLog("New ini:" + WebAccessIP2.Text);
                }
            }
            else
            {
                EventLog.AddLog(".ini file not exist, create new .ini file. Path: " + sIniFilePath);
                tpc.F_WritePrivateProfileString("ProjectName", "Ground PC or Primary PC", ProjectName.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("ProjectName", "Cloud PC or Backup PC", ProjectName2.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("IP", "Ground PC or Primary PC", WebAccessIP.Text, sIniFilePath);
                tpc.F_WritePrivateProfileString("IP", "Cloud PC or Backup PC", WebAccessIP2.Text, sIniFilePath);
            }
        }

    }
}
