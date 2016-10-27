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

namespace PlugandPlay_DeleteUpdateTagTest_GtoC
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
            EventLog.AddLog("===PlugandPlay_DeleteUpdateTagTest_GtoC start (by iATester)===");
            if (System.IO.File.Exists(sIniFilePath))    // 再load一次
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, ProjectName2.Text, WebAccessIP.Text, WebAccessIP2.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===PlugandPlay_DeleteUpdateTagTest_GtoC end (by iATester)===");

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
            GroundPC_DeleteUpdateTag(sBrowser, sProjectName, sWebAccessIP, sTestLogFolder);

            // Step2: Cloud PC view tag info
            bool bViewandSaveCloudTagInfo = ViewandSaveCloudTagInfo(sBrowser, sProjectName2, sWebAccessIP2, sTestLogFolder);

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

            if (bViewandSaveCloudTagInfo && bSeleniumResult)
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

        private void GroundPC_DeleteUpdateTag(string sBrowser, string sProjectName, string sWebAccessIP, string sTestLogFolder)
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


            // Step1. Delete tag
            EventLog.AddLog("<GroundPC> Delete AT_AI0004/AT_AO0004/AT_DI0004/AT_DO0004/OPCDA_0004/OPCUA_0004/Acc_0004/ConAna_0004/SystemSec_0004 tags start...");
            GroundPC_DeleteTag();

            // Step2. Update tag
            EventLog.AddLog("<GroundPC> Update AT_AI0005/AT_AO0005/AT_DI0005/AT_DO0005/OPCDA_0005/OPCUA_0005/Acc_0005/ConDis_0005/SystemSec_0005 tags start...");
            GroundPC_UpdateTag();   //做這之前要先確認左邊的樹狀圖是"全展開"or"Modsim要展開其它不展開"的狀態

            ReturnSCADAPage(api);
            // Step3. Download project
            EventLog.AddLog("<GroundPC> Download...");
            StartDownload(api, sTestLogFolder);

            api.Quit();
            PrintStep(api, "<GroundPC> Quit browser");
        }

        private bool ViewandSaveCloudTagInfo(string sBrowser, string sProjectName, string sWebAccessIP, string sTestLogFolder)
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
            EventLog.AddLog("<CloudPC> Capture the project manager page");
            api2.LinkWebUI(baseUrl2 + "/broadWeb/bwconfig.asp?username=admin");
            api2.ById("userField").Enter("").Submit().Exe();
            PrintStep(api2, "<CloudPC> Login WebAccess");

            // Configure project by project name
            EventLog.AddLog("<CloudPC> Capture the configure project page");
            api2.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
            PrintStep(api2, "<CloudPC> Configure project");

            
            api2.SwitchToCurWindow(0);
            api2.SwitchToFrame("leftFrame", 0);
            
            api2.ByXpath("//tr[4]/td/a/font").Click();      // Acc_0005_Test
            bool bAccCheck = AnalogTagCheck("Acc_0005_Test");

            api2.ByXpath("//tr[253]/td/a/font").Click();    // AT_AI0005_Test
            bool bAICheck = AnalogTagCheck("AT_AI0005_Test");

            api2.ByXpath("//tr[502]/td/a/font").Click();    // AT_AO0005_Test
            bool bAOCheck = AnalogTagCheck("AT_AO0005_Test");
            
            api2.ByXpath("//tr[751]/td/a/font").Click();    // AT_DI0005_Test
            bool bDICheck = DiscreteTagUpdateCheck("AT_DI0005_Test");

            api2.ByXpath("//tr[1000]/td/a/font").Click();   // AT_DO0005_Test
            bool bDOCheck = DiscreteTagUpdateCheck("AT_DO0005_Test");

            api2.ByXpath("//tr[1635]/td/a/font").Click();   // ConDis_0005_Test
            bool bConDisCheck = DiscreteTagUpdateCheck("ConDis_0005_Test");
            
            api2.ByXpath("//tr[2134]/td/a/font").Click();   // OPCDA_0005_Test
            bool bOPCDACheck = AnalogTagCheck("OPCDA_0005_Test");

            api2.ByXpath("//tr[2383]/td/a/font").Click();   // OPCUA_0005_Test
            bool bOPCUACheck = AnalogTagCheck("OPCUA_0005_Test");

            api2.ByXpath("//tr[2632]/td/a/font").Click();   // SystemSec_0005_Test
            bool bSysCheck = AnalogTagCheck("SystemSec_0005_Test");
            
            api2.Quit();
            PrintStep(api2, "<CloudPC> Quit browser");

            if (bAccCheck && bAICheck && bAOCheck && bDICheck && bDOCheck && bConDisCheck && bOPCDACheck && bOPCUACheck && bSysCheck)
                return true;
            else
                return false;
        }

        private void GroundPC_DeleteTag()
        {
            // Delete AT_AI0004/AT_AO0004/AT_DI0004/AT_DO0004/OPCDA_0004/OPCUA_0004/Acc_0004/ConAna_0004/SystemSec_0004
            EventLog.AddLog("<GroundPC> Delete AT_AI0004/AT_AO0004/AT_DI0004/AT_DO0004 tags");
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
            api.ByXpath("//a[3]/img").Click();  //ModSim

            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("(//input[@name='tsel'])[4]").Click();      //AT_AI0004
            api.ByXpath("(//input[@name='tsel'])[254]").Click();    //AT_AO0004
            api.ByXpath("(//input[@name='tsel'])[504]").Click();    //AT_DI0004
            api.ByXpath("(//input[@name='tsel'])[754]").Click();    //AT_DO0004


            api.ByCss("span.e6").Click();   // delete
            string alertText = api.GetAlartTxt();
            if (alertText == "Are you sure you want to delete these tags?")
                api.Accept();

            PrintStep(api, "<GroundPC> Delete AT_AI0004/AT_AO0004/AT_DI0004/AT_DO0004 tags");

            EventLog.AddLog("<GroundPC> Delete OPCDA_0004 tag");
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
            api.ByXpath("//tr[2]/td/table/tbody/tr/td/a[3]/img").Click();   //OPCDA

            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("(//input[@name='tsel'])[4]").Click();      //OPCDA_0004


            api.ByCss("span.e6").Click();   // delete
            string alertText2 = api.GetAlartTxt();
            if (alertText2 == "Are you sure you want to delete these tags?")
                api.Accept();

            PrintStep(api, "<GroundPC> Delete OPCDA_0004 tag");

            EventLog.AddLog("<GroundPC> Delete OPCUA_0004 tag");
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
            api.ByXpath("//tr[3]/td/table/tbody/tr/td/a[3]/img").Click();   //OPCUA

            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("(//input[@name='tsel'])[4]").Click();      //OPCUA_0004


            api.ByCss("span.e6").Click();   // delete
            string alertText3 = api.GetAlartTxt();
            if (alertText3 == "Are you sure you want to delete these tags?")
                api.Accept();

            PrintStep(api, "<GroundPC> Delete OPCUA_0004 tag");

            EventLog.AddLog("<GroundPC> Delete Acc_0004 tag");
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMainRight.asp?pos=AccList')]").Click();   //Acc Point

            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("(//input[@name='tsel'])[4]").Click();      //Acc_0004


            api.ByXpath("(//a[contains(text(),'Delete')])[4]").Click();   // delete
            string alertText4 = api.GetAlartTxt();
            if (alertText4 == "Delete this Accumulation Point(TagName=Acc_0004), are you sure? ")
                api.Accept();

            PrintStep(api, "<GroundPC> Delete Acc_0004 tag");

            EventLog.AddLog("<GroundPC> Delete ConAna_0004 tag");
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMainRight.asp?pos=ConstList')]").Click();   //Const Point

            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("(//input[@name='tsel'])[4]").Click();      //ConAna_0004


            api.ByXpath("(//a[contains(text(),'Delete')])[4]").Click();   // delete
            string alertText5 = api.GetAlartTxt();
            if (alertText5 == "Delete this Constant Point(TagName=ConAna_0004), are you sure? ")
                api.Accept();

            PrintStep(api, "<GroundPC> Delete ConAna_0004 tag");

            EventLog.AddLog("<GroundPC> Delete SystemSec_0004 tag");
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMainRight.asp?pos=SysList')]").Click();   //System Point

            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("(//input[@name='tsel'])[4]").Click();      //SystemSec_0004


            api.ByXpath("(//a[contains(text(),'Delete')])[4]").Click();   // delete
            string alertText6 = api.GetAlartTxt();
            if (alertText6 == "Delete this System Point(TagName=SystemSec_0004), are you sure? ")
                api.Accept();

            PrintStep(api, "<GroundPC> Delete SystemSec_0004 tag");

        }

        private void GroundPC_UpdateTag()
        {
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
            /*
            int iErrorCode = api.ByCss("a[name=\"tag5\"] > font.e5").Click();   //AT_AI0005     // 不能使用bycss 因為css的值好像是由建立tag的順序來設定 會亂掉
            if (iErrorCode != (int)ErrorCode.SUCCESS)   // 展開左列 ModSim tag list
            {
                api.ByXpath("//td[2]/table/tbody/tr/td/table/tbody/tr/td/a/img").Click();
                Thread.Sleep(2000);
                api.ByCss("a[name=\"tag5\"] > font.e5").Click();
            }

            {
                api.ByXpath("//a[contains(@href, '/broadWeb/tag/TagPg.asp?pos=tag') and contains(@href, 'action=tag_property')]").Click();
                api.ByName("AlarmStatus").SelectTxt("Alarm").Exe();
            }
            Thread.Sleep(1500);
            api.ByCss("a[name=\"tag10\"] > font.e5").Click(); //AT_AO0005
            Thread.Sleep(1500);
            api.ByCss("a[name=\"tag15\"] > font.e5").Click(); //AT_DI0005
            Thread.Sleep(1500);
            api.ByCss("a[name=\"tag20\"] > font.e5").Click(); //AT_DO0005
            Thread.Sleep(1500);
            */
            /*
            int iErrorCode = api.ByXpath("//table[@id='table1']/tbody/tr[2]/td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[5]/td/a/font").Click();   //AT_AI0005
            if (iErrorCode != (int)ErrorCode.SUCCESS)   // 展開左列 ModSim tag list
            {
                api.ByXpath("//td[2]/table/tbody/tr/td/table/tbody/tr/td/a/img").Click();
                Thread.Sleep(2000);
                api.ByXpath(@"//table[@id='table1']/tbody/tr[2]/td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[5]/td/a/font").Click();
            }

            Thread.Sleep(1500);
            api.ByXpath(@"//table[@id='table1']/tbody/tr[2]/td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[255]/td/a/font").Click(); //AT_AO0005
            Thread.Sleep(1500);
            api.ByXpath(@"//table[@id='table1']/tbody/tr[2]/td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[505]/td/a/font").Click(); //AT_DI0005
            Thread.Sleep(1500);
            api.ByXpath(@"//table[@id='table1']/tbody/tr[2]/td[2]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[755]/td/a/font").Click(); //AT_DO0005
            Thread.Sleep(1500);
            */

            int iErrorCode = api.ByXpath("//tr[4]/td/a/font").Click();   //AT_AI0005
            if (iErrorCode != (int)ErrorCode.SUCCESS)   // 展開左列 ModSim tag list
            {
                EventLog.AddLog("<GroundPC> Cannot find AT_AI0005.. expand Modsim tree list");
                api.ByXpath("//td[2]/table/tbody/tr/td/table/tbody/tr/td/a/img").Click();
                Thread.Sleep(2000);

                api.ByXpath("//tr[4]/td/a/font").Click();
                Thread.Sleep(1500);
                AnalogTagUpdateSetting("AT_AI0005");

                api.ByXpath("//tr[253]/td/a/font").Click(); //AT_AO0005
                Thread.Sleep(1500);
                AnalogTagUpdateSetting("AT_AO0005");

                api.ByXpath("//tr[502]/td/a/font").Click(); //AT_DI0005
                Thread.Sleep(1500);
                DiscreteTagUpdateSetting("AT_DI0005");

                api.ByXpath("//tr[751]/td/a/font").Click(); //AT_DO0005
                Thread.Sleep(1500);
                DiscreteTagUpdateSetting("AT_DO0005");
            }
            else
            {
                api.ByXpath("//tr[4]/td/a/font").Click();
                Thread.Sleep(1500);
                AnalogTagUpdateSetting("AT_AI0005");

                api.ByXpath("//tr[253]/td/a/font").Click(); //AT_AO0005
                Thread.Sleep(1500);
                AnalogTagUpdateSetting("AT_AO0005");

                api.ByXpath("//tr[502]/td/a/font").Click(); //AT_DI0005
                Thread.Sleep(1500);
                DiscreteTagUpdateSetting("AT_DI0005");

                api.ByXpath("//tr[751]/td/a/font").Click(); //AT_DO0005
                Thread.Sleep(1500);
                DiscreteTagUpdateSetting("AT_DO0005");
            }

            //int iErrorCode2 = api.ByCss("a[name=\"tag186\"] > font.e5").Click();   //OPCDA_0005 
            int iErrorCode2 = api.ByXpath(@"//table[@id='table1']/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[4]/td/a/font").Click();
            if (iErrorCode2 != (int)ErrorCode.SUCCESS)   // 展開左列 OPCDA tag list
            {
                EventLog.AddLog("<GroundPC> Cannot find OPCDA_0005.. expand OPCDA tree list");
                api.ByXpath("//td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/a/img").ClickAndWait(1500);
                //api.ByCss("a[name=\"tag186\"] > font.e5").Click();
                api.ByXpath(@"//table[@id='table1']/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[4]/td/a/font").Click();
                AnalogTagUpdateSetting("OPCDA_0005");
            }
            else
            {
                AnalogTagUpdateSetting("OPCDA_0005");
            }

            Thread.Sleep(2000);
            //int iErrorCode3 = api.ByCss("a[name=\"tag191\"] > font.e5").Click();   //OPCUA_0005
            int iErrorCode3 = api.ByXpath(@"//table[@id='table1']/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr[4]/td/a/font").Click();
            if (iErrorCode3 != (int)ErrorCode.SUCCESS)   // 展開左列 OPCUA tag list
            {
                EventLog.AddLog("<GroundPC> Cannot find OPCUA_0005.. expand OPCUA tree list");
                api.ByXpath("//tr[3]/td/table/tbody/tr/td/a/img").ClickAndWait(1500);
                //api.ByCss("a[name=\"tag191\"] > font.e5").Click();
                api.ByXpath(@"//table[@id='table1']/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table/tbody/tr/td/table/tbody/tr[4]/td/a/font").Click();
                AnalogTagUpdateSetting("OPCUA_0005");
            }
            else
            {
                AnalogTagUpdateSetting("OPCUA_0005");
            }

            Thread.Sleep(2000);
            api.ByXpath("//table[2]/tbody/tr/td/a/img").ClickAndWait(1500); // 預設是關閉的 要展開他
            int iErrorCode4 = api.ByXpath(@"//table[@id='table1']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td/table/tbody/tr[4]/td/a/font").Click();   //Acc_0005
            if (iErrorCode4 != (int)ErrorCode.SUCCESS)   // 展開左列 Acc Point tag list
            {
                EventLog.AddLog("<GroundPC> Cannot find Acc_0005.. expand Acc Point tree list");
                api.ByXpath("//table[2]/tbody/tr/td/a/img").ClickAndWait(1500);
                api.ByXpath(@"//table[@id='table1']/tbody/tr[2]/td[2]/table[2]/tbody/tr/td/table/tbody/tr[4]/td/a/font").Click();
                AnalogTagUpdateSetting("Acc_0005");
            }
            else
            {
                AnalogTagUpdateSetting("Acc_0005");
            }


            Thread.Sleep(2000);
            api.ByXpath("//table[4]/tbody/tr/td/a/img").ClickAndWait(1500); // 預設是關閉的 要展開他
            int iErrorCode5 = api.ByXpath(@"//table[@id='table1']/tbody/tr[2]/td[2]/table[4]/tbody/tr/td/table/tbody/tr[254]/td/a/font").Click();   //ConDis_0005
            if (iErrorCode5 != (int)ErrorCode.SUCCESS)   // 展開左列 Calc Point tag list
            {
                EventLog.AddLog("<GroundPC> Cannot find ConDis_0005.. expand Const Point tree list");
                api.ByXpath("//table[4]/tbody/tr/td/a/img").ClickAndWait(1500);
                api.ByXpath(@"//table[@id='table1']/tbody/tr[2]/td[2]/table[4]/tbody/tr/td/table/tbody/tr[254]/td/a/font").Click();
                DiscreteTagUpdateSetting("ConDis_0005");
            }
            else
            {
                DiscreteTagUpdateSetting("ConDis_0005");
            }

            Thread.Sleep(2000);
            int iErrorCode6 = api.ByXpath(@"//table[@id='table1']/tbody/tr[2]/td[2]/table[5]/tbody/tr/td/table/tbody/tr[4]/td/a/font").Click();   //SystemSec_0005
            if (iErrorCode6 != (int)ErrorCode.SUCCESS)   // 展開左列 System Point tag list
            {
                EventLog.AddLog("<GroundPC> Cannot find SystemSec_0005.. expand System Point tree list");
                api.ByXpath("//table[5]/tbody/tr/td/a/img").ClickAndWait(1000);
                api.ByXpath(@"//table[@id='table1']/tbody/tr[2]/td[2]/table[5]/tbody/tr/td/table/tbody/tr[4]/td/a/font").Click();
                AnalogTagUpdateSetting("SystemSec_0005");
            }
            else
            {
                AnalogTagUpdateSetting("SystemSec_0005");
            }
            Thread.Sleep(2000);

        }

        private void AnalogTagUpdateSetting(string sTagName)
        {
            EventLog.AddLog("<GroundPC> Update " + sTagName);
            Thread.Sleep(2000);
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/tag/TagPg.asp?pos=tag') and contains(@href, 'action=tag_property')]").Click();
            api.ByName("AlarmStatus").SelectTxt("Alarm").Exe();

            Thread.Sleep(1500);
            api.ByName("Description").Clear();
            api.ByName("Description").Enter("Plug and play update tag test from ground to cloud").Exe();
            api.ByName("HHPriority").SelectVal("8").Exe();
            api.ByName("HHAlarm").Clear();
            api.ByName("HHAlarm").Enter("7").Exe();
            api.ByName("HiPriority").SelectVal("6").Exe();
            api.ByName("HiAlarm").Clear();
            api.ByName("HiAlarm").Enter("5").Exe();
            api.ByName("LoPriority").SelectVal("1").Exe();
            api.ByName("LoAlarm").Clear();
            api.ByName("LoAlarm").Enter("4").Exe();
            api.ByName("LLPriority").SelectVal("2").Exe();
            api.ByName("LLAlarm").Clear();
            api.ByName("LLAlarm").Enter("3").Exe();
            api.ByName("TagName").Clear();
            api.ByName("TagName").Enter(sTagName + "_Test").Submit().Exe();
            PrintStep(api, "<GroundPC> Update " + sTagName);

            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
        }

        private void DiscreteTagUpdateSetting(string sTagName)
        {
            EventLog.AddLog("<GroundPC> Update " + sTagName);
            Thread.Sleep(2000);
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/tag/TagPg.asp?pos=tag') and contains(@href, 'action=tag_property')]").Click();
            api.ByName("AlarmStatus").SelectTxt("Alarm").Exe();

            Thread.Sleep(1500);
            api.ByName("Description").Clear();
            api.ByName("Description").Enter("Plug and play update tag test from ground to cloud").Exe();
            api.ByName("AlarmPriority0").SelectVal("8").Exe();
            api.ByName("DelayTime0").Clear();
            api.ByName("DelayTime0").Enter("0").Exe();
            api.ByName("AlarmPriority1").SelectVal("7").Exe();
            api.ByName("DelayTime1").Clear();
            api.ByName("DelayTime1").Enter("0").Exe();
            api.ByName("AlarmPriority2").SelectVal("6").Exe();
            api.ByName("DelayTime2").Clear();
            api.ByName("DelayTime2").Enter("0").Exe();
            api.ByName("AlarmPriority3").SelectVal("5").Exe();
            api.ByName("DelayTime3").Clear();
            api.ByName("DelayTime3").Enter("0").Exe();
            api.ByName("AlarmPriority4").SelectVal("4").Exe();
            api.ByName("DelayTime4").Clear();
            api.ByName("DelayTime4").Enter("0").Exe();
            api.ByName("AlarmPriority5").SelectVal("3").Exe();
            api.ByName("DelayTime5").Clear();
            api.ByName("DelayTime5").Enter("0").Exe();
            api.ByName("AlarmPriority6").SelectVal("2").Exe();
            api.ByName("DelayTime6").Clear();
            api.ByName("DelayTime6").Enter("0").Exe();
            api.ByName("AlarmPriority7").SelectVal("1").Exe();
            api.ByName("TagName").Clear();
            api.ByName("TagName").Enter(sTagName + "_Test").Submit().Exe();
            PrintStep(api, "<GroundPC> Update " + sTagName);

            api.SwitchToCurWindow(0);
            api.SwitchToFrame("leftFrame", 0);
        }

        private bool AnalogTagCheck(string sTagName)
        {
            EventLog.AddLog("<CloudPC> Update analog tag check: "+ sTagName);
            Thread.Sleep(500);
            api2.SwitchToCurWindow(0);
            api2.SwitchToFrame("rightFrame", 0);
            string sTagChangedName = api2.ByXpath("//tr[2]/td[2]").GetText();       // OPCDA_0005_Test (Tag Name)
            string sDescription = api2.ByXpath("//tr[3]/td[2]").GetText();          // Plug and play update tag test from ground to cloud (Description)
            string sHHPriority = api2.ByXpath("//tr[35]/td[2]/font").GetText();     // 8 (HH Priority)
            string sHHAlarmLimit = api2.ByXpath("//tr[36]/td[2]/font").GetText();   // 7 (HH Alarm Limit)
            string sHighPriority = api2.ByXpath("//tr[37]/td[2]/font").GetText();   // 6 (High Priority)
            string sHighAlarmLimit = api2.ByXpath("//tr[38]/td[2]/font").GetText(); // 5 (High Alarm Limit)
            string sLowPriority = api2.ByXpath("//tr[39]/td[2]/font").GetText();    // 1 (Low Priority)
            string sLowAlarmLimit = api2.ByXpath("//tr[40]/td[2]/font").GetText();  // 4 (Low Alarm Limit)
            string sLLPriority = api2.ByXpath("//tr[41]/td[2]/font").GetText();     // 2 (LL Priority)
            string sLLAlarmLimit = api2.ByXpath("//tr[42]/td[2]/font").GetText();   // 3 (LL Alarm Limit)
            string sHLDb = api2.ByXpath("//tr[43]/td[2]/font").GetText();           // 0 (HL Db)
            api2.SwitchToCurWindow(0);
            api2.SwitchToFrame("leftFrame", 0);

            EventLog.AddLog("<CloudPC> Tag Name = " + sTagChangedName.Trim());
            EventLog.AddLog("<CloudPC> Description = " + sDescription.Trim());
            EventLog.AddLog("<CloudPC> HHPriority = " + sHHPriority.Trim());
            EventLog.AddLog("<CloudPC> HHAlarmLimit = " + sHHAlarmLimit.Trim());
            EventLog.AddLog("<CloudPC> HighPriority = " + sHighPriority.Trim());
            EventLog.AddLog("<CloudPC> HighAlarmLimit = " + sHighAlarmLimit.Trim());
            EventLog.AddLog("<CloudPC> LowPriority = " + sLowPriority.Trim());
            EventLog.AddLog("<CloudPC> LowAlarmLimit = " + sLowAlarmLimit.Trim());
            EventLog.AddLog("<CloudPC> LLPriority = " + sLLPriority.Trim());
            EventLog.AddLog("<CloudPC> LLAlarmLimit = " + sLLAlarmLimit.Trim());
            EventLog.AddLog("<CloudPC> HLDb = " + sHLDb.Trim());

            PrintStep(api2, "<CloudPC> Check" + sTagName + " info");
            if (sTagChangedName.Trim() != sTagName ||
                sDescription.Trim() != "Plug and play update tag test from ground to cloud" ||
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
                EventLog.AddLog("<CloudPC> " + sTagName + " check fail !!!");
                return false;
            }
            else
            {
                EventLog.AddLog("<CloudPC> " + sTagName + " check pass !!!");
                return true;
            }
        }

        private bool DiscreteTagUpdateCheck(string sTagName)
        {
            EventLog.AddLog("<CloudPC> Update discrete tag check: " + sTagName);
            Thread.Sleep(500);
            api2.SwitchToCurWindow(0);
            api2.SwitchToFrame("rightFrame", 0);
            string sdTagChangedName = api2.ByXpath("//tr[2]/td[2]").GetText();       // OPCDA_0005_Test (Tag Name)
            string sdDescription = api2.ByXpath("//tr[3]/td[2]").GetText();          // Plug and play update tag test from ground to cloud (Description)

            string sdState0AlarmPriority = api2.ByXpath("//tr[31]/td[2]/font").GetText();     // 8 (State 0 Alarm Priority)
            string sdAlarmDelayTime0 = api2.ByXpath("//tr[32]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
            string sdState1AlarmPriority = api2.ByXpath("//tr[33]/td[2]/font").GetText();     // 7 (State 1 Alarm Priority)
            string sdAlarmDelayTime1 = api2.ByXpath("//tr[34]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
            string sdState2AlarmPriority = api2.ByXpath("//tr[35]/td[2]/font").GetText();     // 6 (State 2 Alarm Priority)
            string sdAlarmDelayTime2 = api2.ByXpath("//tr[36]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
            string sdState3AlarmPriority = api2.ByXpath("//tr[37]/td[2]/font").GetText();     // 5 (State 3 Alarm Priority)
            string sdAlarmDelayTime3 = api2.ByXpath("//tr[38]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
            string sdState4AlarmPriority = api2.ByXpath("//tr[39]/td[2]/font").GetText();     // 4 (State 4 Alarm Priority)
            string sdAlarmDelayTime4 = api2.ByXpath("//tr[40]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
            string sdState5AlarmPriority = api2.ByXpath("//tr[41]/td[2]/font").GetText();     // 3 (State 5 Alarm Priority)
            string sdAlarmDelayTime5 = api2.ByXpath("//tr[42]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
            string sdState6AlarmPriority = api2.ByXpath("//tr[43]/td[2]/font").GetText();     // 2 (State 6 Alarm Priorityb)
            string sdAlarmDelayTime6 = api2.ByXpath("//tr[44]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
            string sdState7AlarmPriority = api2.ByXpath("//tr[45]/td[2]/font").GetText();     // 1 (State 7 Alarm Priority)
            string sdAlarmDelayTime7 = api2.ByXpath("//tr[46]/td[2]/font").GetText();         // 0 (Alarm Delay Time)
            api2.SwitchToCurWindow(0);
            api2.SwitchToFrame("leftFrame", 0);

            EventLog.AddLog("<CloudPC> Tag Name = " + sdTagChangedName.Trim());
            EventLog.AddLog("<CloudPC> Description = " + sdDescription.Trim());

            EventLog.AddLog("<CloudPC> State 0 AlarmPriority = " + sdState0AlarmPriority.Trim());
            EventLog.AddLog("<CloudPC> AlarmDelayTime 0 = " + sdAlarmDelayTime0.Trim());

            EventLog.AddLog("<CloudPC> State 1 AlarmPriority = " + sdState1AlarmPriority.Trim());
            EventLog.AddLog("<CloudPC> AlarmDelayTime 1 = " + sdAlarmDelayTime1.Trim());

            EventLog.AddLog("<CloudPC> State 2 AlarmPriority = " + sdState2AlarmPriority.Trim());
            EventLog.AddLog("<CloudPC> AlarmDelayTime 2 = " + sdAlarmDelayTime2.Trim());

            EventLog.AddLog("<CloudPC> State 3 AlarmPriority = " + sdState3AlarmPriority.Trim());
            EventLog.AddLog("<CloudPC> AlarmDelayTime 3 = " + sdAlarmDelayTime3.Trim());

            EventLog.AddLog("<CloudPC> State 4 AlarmPriority = " + sdState4AlarmPriority.Trim());
            EventLog.AddLog("<CloudPC> sdAlarmDelayTime 4 = " + sdAlarmDelayTime4.Trim());

            EventLog.AddLog("<CloudPC> State 5 AlarmPriority = " + sdState5AlarmPriority.Trim());
            EventLog.AddLog("<CloudPC> AlarmDelayTime 5 = " + sdAlarmDelayTime5.Trim());

            EventLog.AddLog("<CloudPC> State 6 AlarmPriority = " + sdState6AlarmPriority.Trim());
            EventLog.AddLog("<CloudPC> AlarmDelayTime 6 = " + sdAlarmDelayTime6.Trim());

            EventLog.AddLog("<CloudPC> State 7 AlarmPriority = " + sdState7AlarmPriority.Trim());
            EventLog.AddLog("<CloudPC> AlarmDelayTime 7 = " + sdAlarmDelayTime7.Trim());

            PrintStep(api2, "<CloudPC> Check" + sTagName + " info");
            if (sdTagChangedName.Trim() != sTagName ||
                sdDescription.Trim() != "Plug and play update tag test from ground to cloud" ||
                sdState0AlarmPriority.Trim() != "8" ||
                sdAlarmDelayTime0.Trim()     != "0" ||
                sdState1AlarmPriority.Trim() != "7" ||
                sdAlarmDelayTime1.Trim()     != "0" ||
                sdState2AlarmPriority.Trim() != "6" ||
                sdAlarmDelayTime2.Trim()     != "0" ||
                sdState3AlarmPriority.Trim() != "5" ||
                sdAlarmDelayTime3.Trim()     != "0" ||
                sdState4AlarmPriority.Trim() != "4"||
                sdAlarmDelayTime4.Trim()     != "0"||
                sdState5AlarmPriority.Trim() != "3"||
                sdAlarmDelayTime5.Trim()     != "0"||
                sdState6AlarmPriority.Trim() != "2"||
                sdAlarmDelayTime6.Trim()     != "0"||
                sdState7AlarmPriority.Trim() != "1"||
                sdAlarmDelayTime7.Trim()     != "0")
            {
                EventLog.AddLog("<CloudPC> " + sTagName + " check fail !!!");
                return false;
            }
            else
            {
                EventLog.AddLog("<CloudPC> " + sTagName + " check pass !!!");
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
        }

        private void Start_Click(object sender, EventArgs e)
        {
            long lErrorCode = (long)ErrorCode.SUCCESS;
            EventLog.AddLog("===PlugandPlay_DeleteUpdateTagTest_GtoC start===");
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address(Ground PC)= " + WebAccessIP.Text);
            EventLog.AddLog("WebAccess IP address(Cloud PC)= " + WebAccessIP2.Text);
            lErrorCode = Form1_Load(ProjectName.Text, ProjectName2.Text, WebAccessIP.Text, WebAccessIP2.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===PlugandPlay_DeleteUpdateTagTest_GtoC end===");
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
