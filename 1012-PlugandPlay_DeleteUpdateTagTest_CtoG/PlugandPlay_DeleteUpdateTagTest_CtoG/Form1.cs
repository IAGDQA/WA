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

namespace PlugandPlay_DeleteUpdateTagTest_CtoG
{
    public partial class Form1 : Form, iATester.iCom
    {
        IAdvSeleniumAPI api;
        IAdvSeleniumAPI api2;
        cThirdPartyToolControl tpc = new cThirdPartyToolControl();

        private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
        private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
        internal const int Max_Rows_Val = 65535;
        string baseUrl, baseUrl2;
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
            EventLog.AddLog("===PlugandPlay_DeleteUpdateTagTest_CtoG start (by iATester)===");
            if (System.IO.File.Exists(sIniFilePath))    // 再load一次
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, ProjectName2.Text, WebAccessIP.Text, WebAccessIP2.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===PlugandPlay_DeleteUpdateTagTest_CtoG end (by iATester)===");

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

            // Step1: Ground PC delete and update tag
            EventLog.AddLog("<CloudPC> CloudPC Delete and Update Tag");
            CloudPC_DeleteUpdateTag(sBrowser, sProjectName2, sWebAccessIP2, sTestLogFolder);

            // Step2. View and Save Ground Tag Info
            EventLog.AddLog("<GroundPC> View and Save Ground Tag Info");
            bool bViewandSaveGroundTagInfoResult = ViewandSaveGroundTagInfo(sBrowser, sProjectName, sWebAccessIP, sTestLogFolder);

            bool bSeleniumResult = true;
            int iTotalSeleniumAction = dataGridView1.Rows.Count;
            for (int i = 0; i < iTotalSeleniumAction-1; i++)
            {
                DataGridViewRow row = dataGridView1.Rows[i];
                string sSeleniumResult = row.Cells[2].Value.ToString();
                if(sSeleniumResult != "pass")
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

            if (bViewandSaveGroundTagInfoResult && bSeleniumResult)
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

        private void CloudPC_DeleteUpdateTag(string sBrowser, string sProjectName, string sWebAccessIP, string sTestLogFolder)
        {
            if (sBrowser == "Internet Explorer")
            {
                EventLog.AddLog("<CloudPC> Browser= Internet Explorer");
                api2 = new AdvSeleniumAPI("IE", "");
                System.Threading.Thread.Sleep(1000);
            }
            else if (sBrowser == "Mozilla FireFox")
            {
                EventLog.AddLog("<CloudPC> Browser= Mozilla FireFox");
                api2 = new AdvSeleniumAPI("FireFox", "");
                System.Threading.Thread.Sleep(1000);
            }

            api2.LinkWebUI(baseUrl2 + "/broadWeb/bwconfig.asp?username=admin");
            api2.ById("userField").Enter("").Submit().Exe();
            PrintStep(api2, "<CloudPC> Login WebAccess");

            // Configure project by project name
            api2.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
            PrintStep(api2, "<CloudPC> Configure project");


            // Step1. Update tag
            EventLog.AddLog("<CloudPC> Update AT_AI0007/AT_AO0007/AT_DI0007/AT_DO0007/OPCDA_0007/OPCUA_0007/Acc_0007/ConDis_0007/SystemSec_0007 tags start...");
            CloudPC_UpdateTag();

            // Step2. Delete tag
            EventLog.AddLog("<CloudPC> Delete AT_AI0006/AT_AO0006/AT_DI0006/AT_DO0006/OPCDA_0006/OPCUA_0006/Acc_0006/ConAna_0006/SystemSec_0006 tags start...");
            CloudPC_DeleteTag();

            ReturnSCADAPage(api2);
            // Step3. Download project
            EventLog.AddLog("<CloudPC> Download...");
            StartDownload(api2, sTestLogFolder);

            api2.Quit();
            PrintStep(api2, "<CloudPC> Quit browser");
        }

        private bool ViewandSaveGroundTagInfo(string sBrowser, string sProjectName, string sWebAccessIP, string sTestLogFolder)
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

            // Step1: Check deleted tag info from whhitelist
            bool bCheckDeletedTagInfo = CheckDeletedTagInfo();
            if (bCheckDeletedTagInfo == false)
                EventLog.AddLog(">>Check some deleted tag FAIL!!<<");
            else
                EventLog.AddLog(">>Check deleted tag PASS!!<<");

            // Step2: Check updated tag info
            bool bCheckUpdatedTagInfo = CheckUpdatedTagInfo();
            if (bCheckUpdatedTagInfo == false)
                EventLog.AddLog(">>Check some updated tag FAIL!!<<");
            else
                EventLog.AddLog(">>Check updated tag PASS!!<<");
            //if (bCheckUpdatedTagInfo && bCheckUpdatedTagInfo)
            //{
            //    Result.Text = "PASS!!";
            //    Result.ForeColor = Color.Green;
            //}
            //else
            //{
            //    Result.Text = "FAIL!!";
            //    Result.ForeColor = Color.Red;
            //}

            api.Quit();
            PrintStep(api, "<GroundPC> Quit browser");

            return (bCheckUpdatedTagInfo && bCheckUpdatedTagInfo);
        }

        private void CloudPC_DeleteTag()
        {
            // Delete AT_AI0006/AT_AO0006/AT_DI0006/AT_DO0006/OPCDA_0006/OPCUA_0006/Acc_0006/ConAna_0006/SystemSec_0006
            EventLog.AddLog("<CloudPC> Delete AT_AI0006/AT_AO0006/AT_DI0006/AT_DO0006/OPCDA_0006/OPCUA_0006/Acc_0006/ConAna_0006/SystemSec_0006 tags");
            api2.SwitchToCurWindow(0);
            api2.SwitchToFrame("leftFrame", 0);

            api2.ByXpath("//tr[5]/td/a/font").Click();     //Acc_0006
            Thread.Sleep(1000);
            DeleteTag(api2, "Acc_0006");

            api2.ByXpath("//tr[253]/td/a/font").Click();   //AT_AI0006
            Thread.Sleep(1000);
            DeleteTag(api2, "AT_AI0006");

            api2.ByXpath("//tr[501]/td/a/font").Click();   //AT_AO0006
            Thread.Sleep(1000);
            DeleteTag(api2, "AT_AO0006");

            api2.ByXpath("//tr[749]/td/a/font").Click();  //AT_DI0006
            Thread.Sleep(1000);
            DeleteTag(api2, "AT_DI0006");

            api2.ByXpath("//tr[997]/td/a/font").Click();  //AT_DO0006
            Thread.Sleep(1000);
            DeleteTag(api2, "AT_DO0006");

            api2.ByXpath("//tr[1631]/td/a/font").Click();  //ConDis_0006
            Thread.Sleep(1000);
            DeleteTag(api2, "ConDis_0006");

            api2.ByXpath("//tr[2129]/td/a/font").Click();  //OPCDA_0006
            Thread.Sleep(1000);
            DeleteTag(api2, "OPCDA_0006");

            api2.ByXpath("//tr[2377]/td/a/font").Click();  //OPCUA_0006
            Thread.Sleep(1000);
            DeleteTag(api2, "OPCUA_0006");

            api2.ByXpath("//tr[2625]/td/a/font").Click();  //SystemSec_0006
            Thread.Sleep(1000);
            DeleteTag(api2, "SystemSec_0006");

            PrintStep(api2, "<CloudPC> Delete No0006 tags");
        }

        private void DeleteTag(IAdvSeleniumAPI api, string sTagName)
        {
            EventLog.AddLog("<CloudPC> Delete " + sTagName);
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/tag/deleteTag.asp') and contains(@href, 'action=delete_tag')]").Click();   // delete

            string alertText = api.GetAlartTxt();
            if (alertText == "Delete this tag, are you sure?")
                api.Accept();

            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
        }

        private void CloudPC_UpdateTag()
        {
            api2.SwitchToCurWindow(0);
            api2.SwitchToFrame("leftFrame", 0);

            api2.ByXpath("//tr[6]/td/a/font").Click();     //Acc_0007
            Thread.Sleep(1000);
            AnalogTagUpdateSetting(api2, "Acc_0007");

            api2.ByXpath("//tr[255]/td/a/font").Click();   //AT_AI0007
            Thread.Sleep(1000);
            AnalogTagUpdateSetting(api2, "AT_AI0007");

            api2.ByXpath("//tr[504]/td/a/font").Click();   //AT_AO0007
            Thread.Sleep(1000);
            AnalogTagUpdateSetting(api2, "AT_AO0007");

            api2.ByXpath("//tr[753]/td/a/font").Click();  //AT_DI0007
            Thread.Sleep(1000);
            DiscreteTagUpdateSetting(api2, "AT_DI0007");

            api2.ByXpath("//tr[1002]/td/a/font").Click();  //AT_DO0007
            Thread.Sleep(1000);
            DiscreteTagUpdateSetting(api2, "AT_DO0007");

            api2.ByXpath("//tr[1637]/td/a/font").Click();  //ConDis_0007
            Thread.Sleep(1000);
            DiscreteTagUpdateSetting(api2, "ConDis_0007");

            api2.ByXpath("//tr[2136]/td/a/font").Click();  //OPCDA_0007
            Thread.Sleep(1000);
            AnalogTagUpdateSetting(api2, "OPCDA_0007");

            api2.ByXpath("//tr[2385]/td/a/font").Click();  //OPCUA_0007
            Thread.Sleep(1000);
            AnalogTagUpdateSetting(api2, "OPCUA_0007");

            api2.ByXpath("//tr[2634]/td/a/font").Click();  //SystemSec_0007
            Thread.Sleep(1000);
            AnalogTagUpdateSetting(api2, "SystemSec_0007");
            Thread.Sleep(2000);
        }

        private void AnalogTagUpdateSetting(IAdvSeleniumAPI Ana_api, string sTagName)
        {
            EventLog.AddLog("<CloudPC> Update " + sTagName);
            Thread.Sleep(2000);
            Ana_api.SwitchToCurWindow(0);
            Ana_api.SwitchToFrame("rightFrame", 0);
            Ana_api.ByXpath("//a[contains(@href, '/broadWeb/tag/TagPg.asp?pos=tag') and contains(@href, 'action=tag_property')]").Click();
            Ana_api.ByName("AlarmStatus").SelectTxt("Alarm").Exe();

            Thread.Sleep(1500);
            Ana_api.ByName("HHPriority").SelectVal("8").Exe();
            Ana_api.ByName("HHAlarm").Clear();
            Ana_api.ByName("HHAlarm").Enter("7").Exe();
            Ana_api.ByName("HiPriority").SelectVal("6").Exe();
            Ana_api.ByName("HiAlarm").Clear();
            Ana_api.ByName("HiAlarm").Enter("5").Exe();
            Ana_api.ByName("LoPriority").SelectVal("1").Exe();
            Ana_api.ByName("LoAlarm").Clear();
            Ana_api.ByName("LoAlarm").Enter("4").Exe();
            Ana_api.ByName("LLPriority").SelectVal("2").Exe();
            Ana_api.ByName("LLAlarm").Clear();
            Ana_api.ByName("LLAlarm").Enter("3").Exe();
            Ana_api.ByName("Description").Clear();
            Ana_api.ByName("Description").Enter("Plug and play update tag test from cloud to ground").Submit().Exe();
            //Ana_api.ByName("TagName").Clear();
            //Ana_api.ByName("TagName").Enter(sTagName + "_Test").Submit().Exe();
            PrintStep(Ana_api, "<CloudPC> Update " + sTagName);

            Ana_api.SwitchToCurWindow(0);
            Ana_api.SwitchToFrame("leftFrame", 0);
        }

        private void DiscreteTagUpdateSetting(IAdvSeleniumAPI Dis_api, string sTagName)
        {
            EventLog.AddLog("<CloudPC> Update " + sTagName);
            Thread.Sleep(2000);
            Dis_api.SwitchToCurWindow(0);
            Dis_api.SwitchToFrame("rightFrame", 0);
            Dis_api.ByXpath("//a[contains(@href, '/broadWeb/tag/TagPg.asp?pos=tag') and contains(@href, 'action=tag_property')]").Click();
            Dis_api.ByName("AlarmStatus").SelectTxt("Alarm").Exe();

            Thread.Sleep(1500);
            Dis_api.ByName("AlarmPriority0").SelectVal("8").Exe();
            Dis_api.ByName("DelayTime0").Clear();
            Dis_api.ByName("DelayTime0").Enter("0").Exe();
            Dis_api.ByName("AlarmPriority1").SelectVal("7").Exe();
            Dis_api.ByName("DelayTime1").Clear();
            Dis_api.ByName("DelayTime1").Enter("0").Exe();
            Dis_api.ByName("AlarmPriority2").SelectVal("6").Exe();
            Dis_api.ByName("DelayTime2").Clear();
            Dis_api.ByName("DelayTime2").Enter("0").Exe();
            Dis_api.ByName("AlarmPriority3").SelectVal("5").Exe();
            Dis_api.ByName("DelayTime3").Clear();
            Dis_api.ByName("DelayTime3").Enter("0").Exe();
            Dis_api.ByName("AlarmPriority4").SelectVal("4").Exe();
            Dis_api.ByName("DelayTime4").Clear();
            Dis_api.ByName("DelayTime4").Enter("0").Exe();
            Dis_api.ByName("AlarmPriority5").SelectVal("3").Exe();
            Dis_api.ByName("DelayTime5").Clear();
            Dis_api.ByName("DelayTime5").Enter("0").Exe();
            Dis_api.ByName("AlarmPriority6").SelectVal("2").Exe();
            Dis_api.ByName("DelayTime6").Clear();
            Dis_api.ByName("DelayTime6").Enter("0").Exe();
            Dis_api.ByName("AlarmPriority7").SelectVal("1").Exe();
            Dis_api.ByName("Description").Clear();
            Dis_api.ByName("Description").Enter("Plug and play update tag test from cloud to ground").Submit().Exe();
            //Dis_api.ByName("TagName").Clear();
            //Dis_api.ByName("TagName").Enter(sTagName + "_Test").Submit().Exe();
            PrintStep(Dis_api, "<CloudPC> Update " + sTagName);

            Dis_api.SwitchToCurWindow(0);
            Dis_api.SwitchToFrame("leftFrame", 0);
        }

        private bool CheckDeletedTagInfo()
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/WaCloudWhitelist/CloudWhitelist.asp?')]").Click();

            api.ById("tagTypes").SelectTxt("Port3(tcpip)").Exe();
            Thread.Sleep(2000);
            api.ByCss("img").Click();   // page1
            Thread.Sleep(2000);

            bool bTotalResult = true;

            if (bTotalResult == true)
            {
                //string sAI6 = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[6]/td[2]").GetText(); // 20161121版本測試
                //string sAI6_checkbox = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[6]/td/input").GetAttr("checked");
                string sAI6 = api.ByXpath("//tr[5]/td[2]").GetText();
                string sAI6_checkbox = api.ByXpath("//tr[5]/td/input").GetAttr("checked");
                if (sAI6 == "AT_AI0006" && sAI6_checkbox == "false")
                    bTotalResult = true;
                else
                    bTotalResult = false;
                EventLog.AddLog("<GroundPC> Check deleted tag name = " + sAI6);
                EventLog.AddLog("<GroundPC> Check if checkbox is enable? = " + sAI6_checkbox);
                PrintStep(api, "<GroundPC> Check AT_AI0006 info");
            }
            if (bTotalResult == true)
            {
                //string sAO6 = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[255]/td[2]").GetText();// 20161121版本測試
                //string sAO6_checkbox = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[255]/td/input").GetAttr("checked");
                string sAO6 = api.ByXpath("//tr[254]/td[2]").GetText();
                string sAO6_checkbox = api.ByXpath("//tr[254]/td/input").GetAttr("checked");
                if (sAO6 == "AT_AO0006" && sAO6_checkbox == "false")
                    bTotalResult = true;
                else
                    bTotalResult = false;
                EventLog.AddLog("<GroundPC> Check deleted tag name = " + sAO6);
                EventLog.AddLog("<GroundPC> Check if checkbox is enable? = " + sAO6_checkbox);
                PrintStep(api, "<GroundPC> Check AT_AO0006 info");
            }

            api.ByXpath("//a[contains(text(),'2')]").Click();   // page 2
            Thread.Sleep(2000);
            if (bTotalResult == true)
            {
                //string sDI6 = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[4]/td[2]").GetText();// 20161121版本測試
                //string sDI6_checkbox = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[4]/td/input").GetAttr("checked");
                string sDI6 = api.ByXpath("//tr[3]/td[2]").GetText();
                string sDI6_checkbox = api.ByXpath("//tr[3]/td/input").GetAttr("checked");
                if (sDI6 == "AT_DI0006" && sDI6_checkbox == "false")
                    bTotalResult = true;
                else
                    bTotalResult = false;
                EventLog.AddLog("<GroundPC> Check deleted tag name = " + sDI6);
                EventLog.AddLog("<GroundPC> Check if checkbox is enable? = " + sDI6_checkbox);
                PrintStep(api, "<GroundPC> Check AT_DI0006 info");
            }
            if (bTotalResult == true)
            {
                //string sDO6 = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[253]/td[2]").GetText();// 20161121版本測試
                //string sDO6_checkbox = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[253]/td/input").GetAttr("checked");
                string sDO6 = api.ByXpath("//tr[252]/td[2]").GetText();
                string sDO6_checkbox = api.ByXpath("//tr[252]/td/input").GetAttr("checked");
                if (sDO6 == "AT_DO0006" && sDO6_checkbox == "false")
                    bTotalResult = true;
                else
                    bTotalResult = false;
                EventLog.AddLog("<GroundPC> Check deleted tag name = " + sDO6);
                EventLog.AddLog("<GroundPC> Check if checkbox is enable? = " + sDO6_checkbox);
                PrintStep(api, "<GroundPC> Check AT_DO0006 info");
            }

            api.ById("tagTypes").SelectTxt("Port4(opc)").Exe();
            Thread.Sleep(2000);
            api.ByCss("img").Click();   // page1
            Thread.Sleep(2000);
            if (bTotalResult == true)
            {
                //string sOPCDA6 = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[6]/td[2]").GetText();// 20161121版本測試
                //string sOPCDA6_checkbox = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[6]/td/input").GetAttr("checked");
                string sOPCDA6 = api.ByXpath("//tr[5]/td[2]").GetText();
                string sOPCDA6_checkbox = api.ByXpath("//tr[5]/td/input").GetAttr("checked");
                if (sOPCDA6 == "OPCDA_0006" && sOPCDA6_checkbox == "false")
                    bTotalResult = true;
                else
                    bTotalResult = false;
                EventLog.AddLog("<GroundPC> Check deleted tag name = " + sOPCDA6);
                EventLog.AddLog("<GroundPC> Check if checkbox is enable? = " + sOPCDA6_checkbox);
                PrintStep(api, "<GroundPC> Check OPCDA_0006 info");
            }

            api.ById("tagTypes").SelectTxt("Port5(tcpip)").Exe();
            Thread.Sleep(2000);
            api.ByCss("img").Click();   // page1
            Thread.Sleep(2000);
            if (bTotalResult == true)
            {
                //string sOPCUA6 = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[6]/td[2]").GetText();// 20161121版本測試
                //string sOPCUA6_checkbox = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[6]/td/input").GetAttr("checked");
                string sOPCUA6 = api.ByXpath("//tr[5]/td[2]").GetText();
                string sOPCUA6_checkbox = api.ByXpath("//tr[5]/td/input").GetAttr("checked");
                if (sOPCUA6 == "OPCUA_0006" && sOPCUA6_checkbox == "false")
                    bTotalResult = true;
                else
                    bTotalResult = false;
                EventLog.AddLog("<GroundPC> Check deleted tag name = " + sOPCUA6);
                EventLog.AddLog("<GroundPC> Check if checkbox is enable? = " + sOPCUA6_checkbox);
                PrintStep(api, "<GroundPC> Check OPCUA_0006 info");
            }

            api.ById("tagTypes").SelectTxt("Acc Point").Exe();
            //api.ByCss("img").Click();   // page1
            Thread.Sleep(2000);
            if (bTotalResult == true)
            {
                //string sAcc6 = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[6]/td[2]").GetText();// 20161121版本測試
                //string sAcc6_checkbox = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[6]/td/input").GetAttr("checked");
                string sAcc6 = api.ByXpath("//tr[5]/td[2]").GetText();
                string sAcc6_checkbox = api.ByXpath("//tr[5]/td/input").GetAttr("checked");
                if (sAcc6 == "Acc_0006" && sAcc6_checkbox == "false")
                    bTotalResult = true;
                else
                    bTotalResult = false;
                EventLog.AddLog("<GroundPC> Check deleted tag name = " + sAcc6);
                EventLog.AddLog("<GroundPC> Check if checkbox is enable? = " + sAcc6_checkbox);
                PrintStep(api, "<GroundPC> Check Acc_0006 info");
            }

            api.ById("tagTypes").SelectTxt("Const Point").Exe();
            Thread.Sleep(2000);
            if (bTotalResult == true)
            {
                //string sConDis6 = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[256]/td[2]").GetText();// 20161121版本測試
                //string sConDis6_checkbox = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[256]/td/input").GetAttr("checked");
                string sConDis6 = api.ByXpath("//tr[255]/td[2]").GetText();
                string sConDis6_checkbox = api.ByXpath("//tr[255]/td/input").GetAttr("checked");
                if (sConDis6 == "ConDis_0006" && sConDis6_checkbox == "false")
                    bTotalResult = true;
                else
                    bTotalResult = false;
                EventLog.AddLog("<GroundPC> Check deleted tag name = " + sConDis6);
                EventLog.AddLog("<GroundPC> Check if checkbox is enable? = " + sConDis6_checkbox);
                PrintStep(api, "<GroundPC> Check ConDis_0006 info");
            }

            api.ById("tagTypes").SelectTxt("System Point").Exe();
            Thread.Sleep(2000);
            if (bTotalResult == true)
            {
                //string sSys6 = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[6]/td[2]").GetText();// 20161121版本測試
                //string sSys6_checkbox = api.ByXpath("//table[@id='leftTopTable']/tbody/tr[6]/td/input").GetAttr("checked");
                string sSys6 = api.ByXpath("//tr[5]/td[2]").GetText();
                string sSys6_checkbox = api.ByXpath("//tr[5]/td/input").GetAttr("checked");
                if (sSys6 == "SystemSec_0006" && sSys6_checkbox == "false")
                    bTotalResult = true;
                else
                    bTotalResult = false;
                EventLog.AddLog("<GroundPC> Check deleted tag name = " + sSys6);
                EventLog.AddLog("<GroundPC> Check if checkbox is enable? = " + sSys6_checkbox);
                PrintStep(api, "<GroundPC> Check SystemSec_0006 info");
            }

            return bTotalResult;
        }

        private bool CheckUpdatedTagInfo()
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
            
            api.ByXpath("//tr[6]/td/a/font").Click();      // AT_AI0007
            bool bAICheck = AnalogTagCheck("AT_AI0007");

            api.ByXpath("//tr[255]/td/a/font").Click();    // AT_AO0007
            bool bAOCheck = AnalogTagCheck("AT_AO0007");

            api.ByXpath("//tr[504]/td/a/font").Click();    // AT_DI0007
            bool bDICheck = DiscreteTagUpdateCheck("AT_DI0007");

            api.ByXpath("//tr[753]/td/a/font").Click();    // AT_DO0007
            bool bDOCheck = DiscreteTagUpdateCheck("AT_DO0007");

            api.ByXpath("//tr[2]/td/table/tbody/tr/td/table/tbody/tr[6]/td/a/font").Click();   // OPCDA_0007
            bool bOPCDACheck = AnalogTagCheck("OPCDA_0007");

            api.ByXpath("//tr[3]/td/table/tbody/tr/td/table/tbody/tr[6]/td/a/font").Click();   // OPCUA_0007
            bool bOPCUACheck = AnalogTagCheck("OPCUA_0007");
            
            api.ByXpath("//table[2]/tbody/tr/td/table/tbody/tr[6]/td/a/font").Click();   // Acc_0007
            bool bAccCheck = AnalogTagCheck("Acc_0007");

            api.ByXpath("//table[4]/tbody/tr/td/table/tbody/tr[256]/td/a/font").Click();   // ConDis_0007
            bool bConDisCheck = DiscreteTagUpdateCheck("ConDis_0007");

            api.ByXpath("//table[5]/tbody/tr/td/table/tbody/tr[6]/td/a/font").Click();   // SystemSec_0007
            bool bSysCheck = AnalogTagCheck("SystemSec_0007");

            if (bAccCheck && bAICheck && bAOCheck && bDICheck && bDOCheck && bConDisCheck && bOPCUACheck && bOPCDACheck && bSysCheck)
                return true;
            else
                return false;
        }

        private bool AnalogTagCheck(string sTagName)
        {
            EventLog.AddLog("<GroundPC> Update analog tag check: " + sTagName);
            Thread.Sleep(500);
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            string sTagChangedName, sDescription, sHHPriority, sHHAlarmLimit, sHighPriority, sHighAlarmLimit, sLowPriority, sLowAlarmLimit, sLLPriority, sLLAlarmLimit, sHLDb;
            string sTagType = api.ByXpath("//tr/td[2]").GetText();

            // 只有在地面端才需要這樣分類 雲端全部都是Point  (analog)類型
            if (sTagType == "Accumulation")
            {
                sTagChangedName = api.ByXpath("//tr[2]/td[2]").GetText();       // OPCDA_0005_Test (Tag Name)
                sDescription = api.ByXpath("//tr[3]/td[2]").GetText();          // Plug and play update tag test from ground to cloud (Description)
                sHHPriority = api.ByXpath("//tr[29]/td[2]/font").GetText();     // 8 (HH Priority)
                sHHAlarmLimit = api.ByXpath("//tr[30]/td[2]/font").GetText();   // 7 (HH Alarm Limit)
                sHighPriority = api.ByXpath("//tr[31]/td[2]/font").GetText();   // 6 (High Priority)
                sHighAlarmLimit = api.ByXpath("//tr[32]/td[2]/font").GetText(); // 5 (High Alarm Limit)
                sLowPriority = api.ByXpath("//tr[33]/td[2]/font").GetText();    // 1 (Low Priority)
                sLowAlarmLimit = api.ByXpath("//tr[34]/td[2]/font").GetText();  // 4 (Low Alarm Limit)
                sLLPriority = api.ByXpath("//tr[35]/td[2]/font").GetText();     // 2 (LL Priority)
                sLLAlarmLimit = api.ByXpath("//tr[36]/td[2]/font").GetText();   // 3 (LL Alarm Limit)
                sHLDb = api.ByXpath("//tr[37]/td[2]/font").GetText();           // 0 (HL Db)
            }
            else if (sTagType == "System Point  (analog)")
            {
                sTagChangedName = api.ByXpath("//tr[2]/td[2]").GetText();       // OPCDA_0005_Test (Tag Name)
                sDescription = api.ByXpath("//tr[3]/td[2]").GetText();          // Plug and play update tag test from ground to cloud (Description)
                sHHPriority = api.ByXpath("//tr[25]/td[2]/font").GetText();     // 8 (HH Priority)
                sHHAlarmLimit = api.ByXpath("//tr[26]/td[2]/font").GetText();   // 7 (HH Alarm Limit)
                sHighPriority = api.ByXpath("//tr[27]/td[2]/font").GetText();   // 6 (High Priority)
                sHighAlarmLimit = api.ByXpath("//tr[28]/td[2]/font").GetText(); // 5 (High Alarm Limit)
                sLowPriority = api.ByXpath("//tr[29]/td[2]/font").GetText();    // 1 (Low Priority)
                sLowAlarmLimit = api.ByXpath("//tr[30]/td[2]/font").GetText();  // 4 (Low Alarm Limit)
                sLLPriority = api.ByXpath("//tr[31]/td[2]/font").GetText();     // 2 (LL Priority)
                sLLAlarmLimit = api.ByXpath("//tr[32]/td[2]/font").GetText();   // 3 (LL Alarm Limit)
                sHLDb = api.ByXpath("//tr[33]/td[2]/font").GetText();           // 0 (HL Db)
            }
            else    // Point  (analog)
            {
                sTagChangedName = api.ByXpath("//tr[2]/td[2]").GetText();       // OPCDA_0005_Test (Tag Name)
                sDescription = api.ByXpath("//tr[3]/td[2]").GetText();          // Plug and play update tag test from ground to cloud (Description)
                sHHPriority = api.ByXpath("//tr[35]/td[2]/font").GetText();     // 8 (HH Priority)
                sHHAlarmLimit = api.ByXpath("//tr[36]/td[2]/font").GetText();   // 7 (HH Alarm Limit)
                sHighPriority = api.ByXpath("//tr[37]/td[2]/font").GetText();   // 6 (High Priority)
                sHighAlarmLimit = api.ByXpath("//tr[38]/td[2]/font").GetText(); // 5 (High Alarm Limit)
                sLowPriority = api.ByXpath("//tr[39]/td[2]/font").GetText();    // 1 (Low Priority)
                sLowAlarmLimit = api.ByXpath("//tr[40]/td[2]/font").GetText();  // 4 (Low Alarm Limit)
                sLLPriority = api.ByXpath("//tr[41]/td[2]/font").GetText();     // 2 (LL Priority)
                sLLAlarmLimit = api.ByXpath("//tr[42]/td[2]/font").GetText();   // 3 (LL Alarm Limit)
                sHLDb = api.ByXpath("//tr[43]/td[2]/font").GetText();           // 0 (HL Db)
            }
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);

            EventLog.AddLog("<GroundPC> Tag Name = " + sTagChangedName.Trim());
            EventLog.AddLog("<GroundPC> Description = " + sDescription.Trim());
            EventLog.AddLog("<GroundPC> HHPriority = " + sHHPriority.Trim());
            EventLog.AddLog("<GroundPC> HHAlarmLimit = " + sHHAlarmLimit.Trim());
            EventLog.AddLog("<GroundPC> HighPriority = " + sHighPriority.Trim());
            EventLog.AddLog("<GroundPC> HighAlarmLimit = " + sHighAlarmLimit.Trim());
            EventLog.AddLog("<GroundPC> LowPriority = " + sLowPriority.Trim());
            EventLog.AddLog("<GroundPC> LowAlarmLimit = " + sLowAlarmLimit.Trim());
            EventLog.AddLog("<GroundPC> LLPriority = " + sLLPriority.Trim());
            EventLog.AddLog("<GroundPC> LLAlarmLimit = " + sLLAlarmLimit.Trim());
            EventLog.AddLog("<GroundPC> HLDb = " + sHLDb.Trim());

            PrintStep(api, "<GroundPC> Check" + sTagName + " info");
            if (sTagChangedName.Trim() != sTagName ||
                sDescription.Trim() != "Plug and play update tag test from cloud to ground" ||
                sHHPriority.Trim() != "8" ||
                sHHAlarmLimit.Trim() != "7" ||
                sHighPriority.Trim() != "6" ||
                sHighAlarmLimit.Trim() != "5" ||
                sLowPriority.Trim() != "1" ||
                sLowAlarmLimit.Trim() != "4" ||
                sLLPriority.Trim() != "2" ||
                sLLAlarmLimit.Trim() != "3" ||
                sHLDb.Trim() != "0")
            {
                EventLog.AddLog("<GroundPC> " + sTagName + " check fail !!!");
                return false;
            }
            else
            {
                EventLog.AddLog("<GroundPC> " + sTagName + " check pass !!!");
                return true;
            }
        }

        private bool DiscreteTagUpdateCheck(string sTagName)
        {
            EventLog.AddLog("<GroundPC> Update discrete tag check: " + sTagName);
            Thread.Sleep(500);
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            string sdTagChangedName, sdDescription;
            string sdState0AlarmPriority, sdState1AlarmPriority, sdState2AlarmPriority, sdState3AlarmPriority, sdState4AlarmPriority, sdState5AlarmPriority, sdState6AlarmPriority, sdState7AlarmPriority;
            string sdAlarmDelayTime0, sdAlarmDelayTime1, sdAlarmDelayTime2, sdAlarmDelayTime3, sdAlarmDelayTime4, sdAlarmDelayTime5, sdAlarmDelayTime6, sdAlarmDelayTime7;
            string sTagType = api.ByXpath("//tr/td[2]").GetText();

            if (sTagType == "Constant  (discrete)")
            {
                sdTagChangedName = api.ByXpath("//tr[2]/td[2]").GetText();       // OPCDA_0005_Test (Tag Name)
                sdDescription = api.ByXpath("//tr[3]/td[2]").GetText();          // Plug and play update tag test from ground to cloud (Description)
                sdState0AlarmPriority = api.ByXpath("//tr[26]/td[2]/font").GetText();     // 8 (State 0 Alarm Priority)
                sdAlarmDelayTime0 = api.ByXpath("//tr[27]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
                sdState1AlarmPriority = api.ByXpath("//tr[28]/td[2]/font").GetText();     // 7 (State 1 Alarm Priority)
                sdAlarmDelayTime1 = api.ByXpath("//tr[29]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
                sdState2AlarmPriority = api.ByXpath("//tr[30]/td[2]/font").GetText();     // 6 (State 2 Alarm Priority)
                sdAlarmDelayTime2 = api.ByXpath("//tr[31]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
                sdState3AlarmPriority = api.ByXpath("//tr[32]/td[2]/font").GetText();     // 5 (State 3 Alarm Priority)
                sdAlarmDelayTime3 = api.ByXpath("//tr[33]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
                sdState4AlarmPriority = api.ByXpath("//tr[34]/td[2]/font").GetText();     // 4 (State 4 Alarm Priority)
                sdAlarmDelayTime4 = api.ByXpath("//tr[35]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
                sdState5AlarmPriority = api.ByXpath("//tr[36]/td[2]/font").GetText();     // 3 (State 5 Alarm Priority)
                sdAlarmDelayTime5 = api.ByXpath("//tr[37]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
                sdState6AlarmPriority = api.ByXpath("//tr[38]/td[2]/font").GetText();     // 2 (State 6 Alarm Priorityb)
                sdAlarmDelayTime6 = api.ByXpath("//tr[39]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
                sdState7AlarmPriority = api.ByXpath("//tr[40]/td[2]/font").GetText();     // 1 (State 7 Alarm Priority)
                sdAlarmDelayTime7 = api.ByXpath("//tr[41]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
            }
            else
            {
                sdTagChangedName = api.ByXpath("//tr[2]/td[2]").GetText();       // OPCDA_0005_Test (Tag Name)
                sdDescription = api.ByXpath("//tr[3]/td[2]").GetText();          // Plug and play update tag test from ground to cloud (Description)
                sdState0AlarmPriority = api.ByXpath("//tr[31]/td[2]/font").GetText();     // 8 (State 0 Alarm Priority)
                sdAlarmDelayTime0 = api.ByXpath("//tr[32]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
                sdState1AlarmPriority = api.ByXpath("//tr[33]/td[2]/font").GetText();     // 7 (State 1 Alarm Priority)
                sdAlarmDelayTime1 = api.ByXpath("//tr[34]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
                sdState2AlarmPriority = api.ByXpath("//tr[35]/td[2]/font").GetText();     // 6 (State 2 Alarm Priority)
                sdAlarmDelayTime2 = api.ByXpath("//tr[36]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
                sdState3AlarmPriority = api.ByXpath("//tr[37]/td[2]/font").GetText();     // 5 (State 3 Alarm Priority)
                sdAlarmDelayTime3 = api.ByXpath("//tr[38]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
                sdState4AlarmPriority = api.ByXpath("//tr[39]/td[2]/font").GetText();     // 4 (State 4 Alarm Priority)
                sdAlarmDelayTime4 = api.ByXpath("//tr[40]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
                sdState5AlarmPriority = api.ByXpath("//tr[41]/td[2]/font").GetText();     // 3 (State 5 Alarm Priority)
                sdAlarmDelayTime5 = api.ByXpath("//tr[42]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
                sdState6AlarmPriority = api.ByXpath("//tr[43]/td[2]/font").GetText();     // 2 (State 6 Alarm Priorityb)
                sdAlarmDelayTime6 = api.ByXpath("//tr[44]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
                sdState7AlarmPriority = api.ByXpath("//tr[45]/td[2]/font").GetText();     // 1 (State 7 Alarm Priority)
                sdAlarmDelayTime7 = api.ByXpath("//tr[46]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
            }

            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);

            EventLog.AddLog("<GroundPC> Tag Name = " + sdTagChangedName.Trim());
            EventLog.AddLog("<GroundPC> Description = " + sdDescription.Trim());

            EventLog.AddLog("<GroundPC> State 0 AlarmPriority = " + sdState0AlarmPriority.Trim());
            EventLog.AddLog("<GroundPC> AlarmDelayTime 0 = " + sdAlarmDelayTime0.Trim());

            EventLog.AddLog("<GroundPC> State 1 AlarmPriority = " + sdState1AlarmPriority.Trim());
            EventLog.AddLog("<GroundPC> AlarmDelayTime 1 = " + sdAlarmDelayTime1.Trim());

            EventLog.AddLog("<GroundPC> State 2 AlarmPriority = " + sdState2AlarmPriority.Trim());
            EventLog.AddLog("<GroundPC> AlarmDelayTime 2 = " + sdAlarmDelayTime2.Trim());

            EventLog.AddLog("<GroundPC> State 3 AlarmPriority = " + sdState3AlarmPriority.Trim());
            EventLog.AddLog("<GroundPC> AlarmDelayTime 3 = " + sdAlarmDelayTime3.Trim());

            EventLog.AddLog("<GroundPC> State 4 AlarmPriority = " + sdState4AlarmPriority.Trim());
            EventLog.AddLog("<GroundPC> sdAlarmDelayTime 4 = " + sdAlarmDelayTime4.Trim());

            EventLog.AddLog("<GroundPC> State 5 AlarmPriority = " + sdState5AlarmPriority.Trim());
            EventLog.AddLog("<GroundPC> AlarmDelayTime 5 = " + sdAlarmDelayTime5.Trim());

            EventLog.AddLog("<GroundPC> State 6 AlarmPriority = " + sdState6AlarmPriority.Trim());
            EventLog.AddLog("<GroundPC> AlarmDelayTime 6 = " + sdAlarmDelayTime6.Trim());

            EventLog.AddLog("<GroundPC> State 7 AlarmPriority = " + sdState7AlarmPriority.Trim());
            EventLog.AddLog("<GroundPC> AlarmDelayTime 7 = " + sdAlarmDelayTime7.Trim());

            PrintStep(api, "<GroundPC> Check" + sTagName + " info");
            if (sdTagChangedName.Trim() != sTagName ||
                sdDescription.Trim() != "Plug and play update tag test from cloud to ground" ||
                sdState0AlarmPriority.Trim() != "8" ||
                sdAlarmDelayTime0.Trim() != "0" ||
                sdState1AlarmPriority.Trim() != "7" ||
                sdAlarmDelayTime1.Trim() != "0" ||
                sdState2AlarmPriority.Trim() != "6" ||
                sdAlarmDelayTime2.Trim() != "0" ||
                sdState3AlarmPriority.Trim() != "5" ||
                sdAlarmDelayTime3.Trim() != "0" ||
                sdState4AlarmPriority.Trim() != "4" ||
                sdAlarmDelayTime4.Trim() != "0" ||
                sdState5AlarmPriority.Trim() != "3" ||
                sdAlarmDelayTime5.Trim() != "0" ||
                sdState6AlarmPriority.Trim() != "2" ||
                sdAlarmDelayTime6.Trim() != "0" ||
                sdState7AlarmPriority.Trim() != "1" ||
                sdAlarmDelayTime7.Trim() != "0")
            {
                EventLog.AddLog("<GroundPC> " + sTagName + " check fail !!!");
                return false;
            }
            else
            {
                EventLog.AddLog("<GroundPC> " + sTagName + " check pass !!!");
                return true;
            }
        }

        private void StartDownload(IAdvSeleniumAPI api, string sTestLogFolder)
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//tr[2]/td/a[3]/font").Click();    // "Download" click
            Thread.Sleep(2000);
            EventLog.AddLog("Find pop up download window handle");
            string main; object subobj;                     // Find pop up download window handle
            api.GetWinHandle(out main, out subobj);
            IEnumerator<String> windowIterator = (IEnumerator<String>)subobj;

            List<string> items = new List<string>();
            while (windowIterator.MoveNext())
                items.Add(windowIterator.Current);

            EventLog.AddLog("Main window handle= " + main);
            EventLog.AddLog("Window handle list items[0]= " + items[0]);
            EventLog.AddLog("Window handle list items[1]= " + items[1]);
            if (main != items[1])
            {
                EventLog.AddLog("Switch to items[1]");
                api.SwitchToWinHandle(items[1]);
            }
            else
            {
                EventLog.AddLog("Switch to items[0]");
                api.SwitchToWinHandle(items[0]);
            }
            api.ByName("submit").Enter("").Submit().Exe();

            EventLog.AddLog("Start to download and wait 80 seconds...");
            Thread.Sleep(80000);    // Wait 80s for Download finish
            EventLog.AddLog("It's been wait 80 seconds");
            PrintScreen("Download result", sTestLogFolder);
            api.Close();
            EventLog.AddLog("Close download window and switch to main window");
            api.SwitchToWinHandle(main);

            PrintStep(api, "Download");
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
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMainRight.asp') and contains(@href, 'name=CTestSCADA')]").Click();    //因為在cloud 要改成CTestSCADA
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
            long lErrorCode = (long)ErrorCode.SUCCESS;
            EventLog.AddLog("===PlugandPlay_DeleteUpdateTagTest_CtoG start===");
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address(Ground PC)= " + WebAccessIP.Text);
            EventLog.AddLog("WebAccess IP address(Cloud PC)= " + WebAccessIP2.Text);
            lErrorCode = Form1_Load(ProjectName.Text, ProjectName2.Text, WebAccessIP.Text, WebAccessIP2.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===PlugandPlay_DeleteUpdateTagTest_CtoG end===");
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