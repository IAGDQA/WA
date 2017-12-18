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
using CommonFunction;

namespace PlugandPlay_DeleteProjectTest_CtoG
{
    public partial class Form1 : Form, iATester.iCom
    {
        IAdvSeleniumAPI api;
        IAdvSeleniumAPI api2;
        cThirdPartyToolControl tpc = new cThirdPartyToolControl();
        cWACommonFunction wacf = new cWACommonFunction();
        cEventLog EventLog = new cEventLog();

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
            EventLog.AddLog("===PlugandPlay_DeleteProjectTest_CtoG start (by iATester)===");
            if (System.IO.File.Exists(sIniFilePath))    // 再load一次
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, ProjectName2.Text, WebAccessIP.Text, WebAccessIP2.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===PlugandPlay_DeleteProjectTest_CtoG end (by iATester)===");

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

            // Step1: Cloud PC delete project
            CloudPC_DeleteProject(sBrowser, sProjectName2, sWebAccessIP2, sTestLogFolder);

            // Step2: Ground PC view white list info
            //ViewandSaveGroundWhiteListInfo(sBrowser, sProjectName, sWebAccessIP, sTestLogFolder);
            bool bCheckDeletedTagInfoResult = CheckDeletedTagInfo(sBrowser, sProjectName, sWebAccessIP, sTestLogFolder);

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

            if (bSeleniumResult && bCheckDeletedTagInfoResult)
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

        private void CloudPC_DeleteProject(string sBrowser, string sProjectName, string sWebAccessIP, string sTestLogFolder)
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

            EventLog.AddLog("<CloudPC> Delete " + sProjectName + " project.");
            api2.ByXpath("//a[contains(@href, '/broadWeb/project/deleteProject.asp?') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();

            // Confirm to delete porject
            string alertText = api2.GetAlartTxt();
            //if (alertText == "Delete this project (" + sProjectName + "), are you sure?")
            //    api2.Accept();
            switch (slanguage)
            {
                case "ENG":
                    if (alertText == "Delete this project (" + sProjectName + "), are you sure?")
                        api2.Accept();
                    break;
                case "CHT":
                    if (alertText == "您確定要刪除這個工程(" + sProjectName + ")?")
                        api2.Accept();
                    break;
                case "CHS":
                    if (alertText == "您肯定要删除工程(" + sProjectName + ")吗?")
                        api2.Accept();
                    break;
                case "JPN":
                    if (alertText == "このﾌﾟﾛｼﾞｪｸﾄ (" + sProjectName + ") を削除してもよろしいですか?")
                        api2.Accept();
                    break;
                case "KRN":
                    if (alertText == "이 프로젝트(" + sProjectName + ")를 삭제합니다. 계속하시겠습니까?")
                        api2.Accept();
                    break;
                case "FRN":
                    if (alertText == "Supprimer ce projet (" + sProjectName + "), êtes-vous sûr ?")
                        api2.Accept();
                    break;

                default:
                    if (alertText == "Delete this project (" + sProjectName + "), are you sure?")
                        api2.Accept();
                    break;
            }
            PrintStep(api2, "Delete "+ sProjectName + "Node");

            Thread.Sleep(10000);

            api2.Quit();
            PrintStep(api2, "<CloudPC> Quit browser");
        }

        private bool CheckDeletedTagInfo(string sBrowser, string sProjectName, string sWebAccessIP, string sTestLogFolder)
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
            EventLog.AddLog("<GroundPC> Capture the project manager page");
            api.LinkWebUI(baseUrl + "/broadWeb/bwconfig.asp?username=admin");
            api.ById("userField").Enter("").Submit().Exe();
            PrintStep(api, "<GroundPC> Login WebAccess");

            // Configure project by project name
            EventLog.AddLog("<GroundPC> Capture the configure project page");
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
            PrintStep(api, "<GroundPC> Configure project");

            EventLog.AddLog("<GroundPC> Cloud White list setting");
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/WaCloudWhitelist/CloudWhitelist.asp?')]").Click();
            PrintStep(api, "<GroundPC> Check CloudWhitelist");

            //// AI AO DI DO ////
            EventLog.AddLog("Check AI AO DI DO deleted tag..");
            switch (slanguage)
            {
                case "ENG":
                    api.ById("tagTypes").SelectTxt("Port3(tcpip)").Exe();
                    break;
                case "CHT":
                    api.ById("tagTypes").SelectTxt("通信埠3(tcpip)").Exe();
                    break;
                case "CHS":
                    api.ById("tagTypes").SelectTxt("通讯端口3(tcpip)").Exe();
                    break;
                case "JPN":
                    api.ById("tagTypes").SelectTxt("Port3(tcpip)").Exe();
                    break;
                case "KRN":
                    api.ById("tagTypes").SelectTxt("포트3(tcpip)").Exe();
                    break;
                case "FRN":
                    api.ById("tagTypes").SelectTxt("Port3(tcpip)").Exe();
                    break;

                default:
                    api.ById("tagTypes").SelectTxt("Port3(tcpip)").Exe();
                    break;
            }
            Thread.Sleep(2000);
            api.ById("SubContent").Click();   // page1
            Thread.Sleep(2000);

            bool bTotalResult = true;

            for (int i =1; i<=500; i++)
            {
                string sCheckbox = api.ByXpath(string.Format("//tr[{0}]/td[1]/input[1]", i)).GetAttr("checked");

                if (sCheckbox == "true")
                {
                    string sTagNmae = api.ByXpath(string.Format("//tr[{0}]/td[2]", i)).GetText();
                    EventLog.AddLog("<GroundPC> The deleted tag <" + sTagNmae + "> is not disable in cloud white list. Test fail!!");
                    bTotalResult = false;
                }
            }
            PrintStep(api, "<GroundPC> Check AI AO tag info");

            api.ByXpath("//a[contains(text(),'2')]").Click();   // page 2
            Thread.Sleep(2000);
            for (int i = 1; i <= 496; i++)  // 因為被前面PlugandPlay_DeleteUpdateTagTest_GtoC刪除4個點AT_AI0004/AT_AO0004/AT_DI0004/AT_DO0004 所以剩496個tag
            {                               //所以剩496個tag
                string sCheckbox = api.ByXpath(string.Format("//tr[{0}]/td[1]/input[1]", i)).GetAttr("checked");

                if (sCheckbox == "true")
                {
                    string sTagNmae = api.ByXpath(string.Format("//tr[{0}]/td[2]", i)).GetText();
                    EventLog.AddLog("<GroundPC> The deleted tag <" + sTagNmae + "> is not disable in cloud white list. Test fail!!");
                    bTotalResult = false;
                }
            }
            PrintStep(api, "<GroundPC> Check DI DO tag info");

            //// OPCDA ////
            EventLog.AddLog("Check OPCDA deleted tag..");
            switch (slanguage)
            {
                case "ENG":
                    api.ById("tagTypes").SelectTxt("Port4(opc)").Exe();
                    break;
                case "CHT":
                    api.ById("tagTypes").SelectTxt("通信埠4(opc)").Exe();
                    break;
                case "CHS":
                    api.ById("tagTypes").SelectTxt("通讯端口4(opc)").Exe();
                    break;
                case "JPN":
                    api.ById("tagTypes").SelectTxt("Port4(opc)").Exe();
                    break;
                case "KRN":
                    api.ById("tagTypes").SelectTxt("포트4(opc)").Exe();
                    break;
                case "FRN":
                    api.ById("tagTypes").SelectTxt("Port4(opc)").Exe();
                    break;

                default:
                    api.ById("tagTypes").SelectTxt("Port4(opc)").Exe();
                    break;
            }
            Thread.Sleep(2000);
            api.ById("SubContent").Click();   // page1
            Thread.Sleep(2000);
            for (int i = 1; i <= 249; i++)  // 因為被前面PlugandPlay_DeleteUpdateTagTest_GtoC刪除1個點OPCDA_0004 所以剩249個tag
            {
                string sCheckbox = api.ByXpath(string.Format("//tr[{0}]/td[1]/input[1]", i)).GetAttr("checked");

                if (sCheckbox == "true")
                {
                    string sTagNmae = api.ByXpath(string.Format("//tr[{0}]/td[2]", i)).GetText();
                    EventLog.AddLog("<GroundPC> The deleted tag <" + sTagNmae + "> is not disable in cloud white list. Test fail!!");
                    bTotalResult = false;
                }
            }
            PrintStep(api, "<GroundPC> Check OPCDA tag info");

            //// OPCUA ////
            EventLog.AddLog("Check OPCUA deleted tag..");
            switch (slanguage)
            {
                case "ENG":
                    api.ById("tagTypes").SelectTxt("Port5(tcpip)").Exe();
                    break;
                case "CHT":
                    api.ById("tagTypes").SelectTxt("通信埠5(tcpip)").Exe();
                    break;
                case "CHS":
                    api.ById("tagTypes").SelectTxt("通讯端口5(tcpip)").Exe();
                    break;
                case "JPN":
                    api.ById("tagTypes").SelectTxt("Port5(tcpip)").Exe();
                    break;
                case "KRN":
                    api.ById("tagTypes").SelectTxt("포트5(tcpip)").Exe();
                    break;
                case "FRN":
                    api.ById("tagTypes").SelectTxt("Port5(tcpip)").Exe();
                    break;

                default:
                    api.ById("tagTypes").SelectTxt("Port5(tcpip)").Exe();
                    break;
            }
            Thread.Sleep(2000);
            api.ById("SubContent").Click();   // page1
            Thread.Sleep(2000);
            for (int i = 1; i <= 249; i++)  // 因為被前面PlugandPlay_DeleteUpdateTagTest_GtoC刪除1個點OPCUA_0004 所以剩249個tag
            {
                string sCheckbox = api.ByXpath(string.Format("//tr[{0}]/td[1]/input[1]", i)).GetAttr("checked");

                if (sCheckbox == "true")
                {
                    string sTagNmae = api.ByXpath(string.Format("//tr[{0}]/td[2]", i)).GetText();
                    EventLog.AddLog("<GroundPC> The deleted tag <" + sTagNmae + "> is not disable in cloud white list. Test fail!!");
                    bTotalResult = false;
                }
            }
            PrintStep(api, "<GroundPC> Check OPCUA tag info");

            //// Acc ////
            EventLog.AddLog("Check Acc deleted tag..");
            switch (slanguage)
            {
                case "ENG":
                    api.ById("tagTypes").SelectTxt("Acc Point").Exe();
                    break;
                case "CHT":
                    api.ById("tagTypes").SelectTxt("累算點").Exe();
                    break;
                case "CHS":
                    api.ById("tagTypes").SelectTxt("累算点").Exe();
                    break;
                case "JPN":
                    api.ById("tagTypes").SelectTxt("Acc Point").Exe();
                    break;
                case "KRN":
                    api.ById("tagTypes").SelectTxt("누적 포인트").Exe();
                    break;
                case "FRN":
                    api.ById("tagTypes").SelectTxt("Point d'accumul.").Exe();
                    break;

                default:
                    api.ById("tagTypes").SelectTxt("Acc Point").Exe();
                    break;
            }
            Thread.Sleep(2000);
            for (int i = 1; i <= 249; i++)  // 因為被前面PlugandPlay_DeleteUpdateTagTest_GtoC刪除1個點Acc_0004 所以剩249個tag
            {
                string sCheckbox = api.ByXpath(string.Format("//tr[{0}]/td[1]/input[1]", i)).GetAttr("checked");

                if (sCheckbox == "true")
                {
                    string sTagNmae = api.ByXpath(string.Format("//tr[{0}]/td[2]", i)).GetText();
                    EventLog.AddLog("<GroundPC> The deleted tag <" + sTagNmae + "> is not disable in cloud white list. Test fail!!");
                    bTotalResult = false;
                }
            }
            PrintStep(api, "<GroundPC> Check Acc tag info");

            //// Calc ////
            EventLog.AddLog("Check Calculate deleted tag..");
            switch (slanguage)
            {
                case "ENG":
                    api.ById("tagTypes").SelectTxt("Calc Point").Exe();
                    break;
                case "CHT":
                    api.ById("tagTypes").SelectTxt("計算點").Exe();
                    break;
                case "CHS":
                    api.ById("tagTypes").SelectTxt("计算点").Exe();
                    break;
                case "JPN":
                    api.ById("tagTypes").SelectTxt("Calc Point").Exe();
                    break;
                case "KRN":
                    api.ById("tagTypes").SelectTxt("산출 포인트").Exe();
                    break;
                case "FRN":
                    api.ById("tagTypes").SelectTxt("Point calc.").Exe();
                    break;

                default:
                    api.ById("tagTypes").SelectTxt("Calc Point").Exe();
                    break;
            }
            Thread.Sleep(2000);
            for (int i = 1; i <= 136; i++)
            {
                string sCheckbox = api.ByXpath(string.Format("//tr[{0}]/td[1]/input[1]", i)).GetAttr("checked");

                if (sCheckbox == "true")
                {
                    string sTagNmae = api.ByXpath(string.Format("//tr[{0}]/td[2]", i)).GetText();
                    EventLog.AddLog("<GroundPC> The deleted tag <" + sTagNmae + "> is not disable in cloud white list. Test fail!!");
                    bTotalResult = false;
                }
            }
            PrintStep(api, "<GroundPC> Check Calculate tag info");

            //// Const ////
            EventLog.AddLog("Check Constant deleted tag..");
            switch (slanguage)
            {
                case "ENG":
                    api.ById("tagTypes").SelectTxt("Const Point").Exe();
                    break;
                case "CHT":
                    api.ById("tagTypes").SelectTxt("常數點").Exe();
                    break;
                case "CHS":
                    api.ById("tagTypes").SelectTxt("常数点").Exe();
                    break;
                case "JPN":
                    api.ById("tagTypes").SelectTxt("Const Point").Exe();
                    break;
                case "KRN":
                    api.ById("tagTypes").SelectTxt("상수 포인트").Exe();
                    break;
                case "FRN":
                    api.ById("tagTypes").SelectTxt("Point const.").Exe();
                    break;

                default:
                    api.ById("tagTypes").SelectTxt("Const Point").Exe();
                    break;
            }
            Thread.Sleep(2000);
            for (int i = 1; i <= 500; i++)
            {
                string sCheckbox = api.ByXpath(string.Format("//tr[{0}]/td[1]/input[1]", i)).GetAttr("checked");

                if (sCheckbox == "true")
                {
                    string sTagNmae = api.ByXpath(string.Format("//tr[{0}]/td[2]", i)).GetText();
                    EventLog.AddLog("<GroundPC> The deleted tag <" + sTagNmae + "> is not disable in cloud white list. Test fail!!");
                    bTotalResult = false;
                }
            }
            PrintStep(api, "<GroundPC> Check Const tag info(Page1)");

            api.ByXpath("//a[contains(text(),'2')]").Click();   // page 2
            Thread.Sleep(2000);
            for (int i = 1; i <= 249; i++)  // 因為被前面PlugandPlay_DeleteUpdateTagTest_GtoC刪除1個點ConAna_0004 所以剩249個tag
            {
                string sCheckbox = api.ByXpath(string.Format("//tr[{0}]/td[1]/input[1]", i)).GetAttr("checked");

                if (sCheckbox == "true")
                {
                    string sTagNmae = api.ByXpath(string.Format("//tr[{0}]/td[2]", i)).GetText();
                    EventLog.AddLog("<GroundPC> The deleted tag <" + sTagNmae + "> is not disable in cloud white list. Test fail!!");
                    bTotalResult = false;
                }
            }
            PrintStep(api, "<GroundPC> Check Const tag info(Page2)");

            //// System ////
            EventLog.AddLog("Check System deleted tag..");
            switch (slanguage)
            {
                case "ENG":
                    api.ById("tagTypes").SelectTxt("System Point").Exe();
                    break;
                case "CHT":
                    api.ById("tagTypes").SelectTxt("系統點").Exe();
                    break;
                case "CHS":
                    api.ById("tagTypes").SelectTxt("系统点").Exe();
                    break;
                case "JPN":
                    api.ById("tagTypes").SelectTxt("System Point").Exe();
                    break;
                case "KRN":
                    api.ById("tagTypes").SelectTxt("시스템 포인트").Exe();
                    break;
                case "FRN":
                    api.ById("tagTypes").SelectTxt("System Point").Exe();
                    break;

                default:
                    api.ById("tagTypes").SelectTxt("System Point").Exe();
                    break;
            }
            Thread.Sleep(2000);
            for (int i = 1; i <= 249; i++)  // 因為被前面PlugandPlay_DeleteUpdateTagTest_GtoC刪除1個點SystemSec_0004 所以剩249個tag
            {
                string sCheckbox = api.ByXpath(string.Format("//tr[{0}]/td[1]/input[1]", i)).GetAttr("checked");

                if (sCheckbox == "true")
                {
                    string sTagNmae = api.ByXpath(string.Format("//tr[{0}]/td[2]", i)).GetText();
                    EventLog.AddLog("<GroundPC> The deleted tag <" + sTagNmae + "> is not disable in cloud white list. Test fail!!");
                    bTotalResult = false;
                }
            }
            PrintStep(api, "<GroundPC> Check System tag info");

            api.Quit();
            PrintStep(api, "<GroundPC> Quit browser");

            return bTotalResult;
        }

        private void ViewandSaveGroundWhiteListInfo(string sBrowser, string sProjectName, string sWebAccessIP, string sTestLogFolder)
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
            EventLog.AddLog("<GroundPC> Capture the project manager page");
            api.LinkWebUI(baseUrl + "/broadWeb/bwconfig.asp?username=admin");
            api.ById("userField").Enter("").Submit().Exe();
            PrintStep(api, "<GroundPC> Login WebAccess");

            // Configure project by project name
            EventLog.AddLog("<GroundPC> Capture the configure project page");
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
            PrintStep(api, "<GroundPC> Configure project");

            EventLog.AddLog("<GroundPC> Cloud White list setting");
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/WaCloudWhitelist/CloudWhitelist.asp?')]").Click();
            //"/broadWeb/WaCloudWhitelist/CloudWhitelist.asp?nid=1&amp;name=TestSCADA"
            
            ////////////////////////////////// View Cloud White list Setting //////////////////////////////////
            {   // AI/AO/DI/DO
                EventLog.AddLog("<GroundPC> Modbus tag setting");
                api.ById("tagTypes").SelectTxt("Port3(tcpip)").Exe();
                Thread.Sleep(2000);
                api.ById("SubContent").Click();   // page1
                Thread.Sleep(2000);
                
                EventLog.PrintScreen("PlugandPlay_DeleteProjectTest_CtoG_ModsimWhiteList_Page1");
                api.ByXpath("//a[contains(text(),'2')]").Click();   // page 2
                Thread.Sleep(2000);
                EventLog.PrintScreen("PlugandPlay_DeleteProjectTest_CtoG_ModsimWhiteList_Page2");
            }

            // Port4(opc)
            {
                EventLog.AddLog("<GroundPC> Port4(opc) setting");
                api.ById("tagTypes").SelectTxt("Port4(opc)").Exe();
                Thread.Sleep(2000);
                api.ById("SubContent").Click();   // page1
                Thread.Sleep(2000);
                EventLog.PrintScreen("PlugandPlay_DeleteProjectTest_CtoG_OPCDAWhiteList");
            }

            // Port5(tcpip)
            {
                EventLog.AddLog("<GroundPC> Port5(tcpip) setting");
                api.ById("tagTypes").SelectTxt("Port5(tcpip)").Exe();
                Thread.Sleep(2000);
                api.ById("SubContent").Click();   // page1
                Thread.Sleep(2000);
                EventLog.PrintScreen("PlugandPlay_DeleteProjectTest_CtoG_OPCUAWhiteList");
            }

            // Acc Point
            {
                EventLog.AddLog("<GroundPC> Acc Point setting");
                api.ById("tagTypes").SelectTxt("Acc Point").Exe();
                //api.ById("SubContent").Click();   // page1
                Thread.Sleep(2000);
                EventLog.PrintScreen("PlugandPlay_DeleteProjectTest_CtoG_AccWhiteList");
            }

            // Calc Point
            {
                EventLog.AddLog("<GroundPC> Calc Point setting");
                api.ById("tagTypes").SelectTxt("Calc Point").Exe();
                //api.ById("SubContent").Click();   // page1
                Thread.Sleep(2000);
                EventLog.PrintScreen("PlugandPlay_DeleteProjectTest_CtoG_CalcWhiteList");
            }

            // Const Point
            {
                EventLog.AddLog("<GroundPC> Const Point setting");
                api.ById("tagTypes").SelectTxt("Const Point").Exe();
                //api.ById("SubContent").Click();   // page1
                Thread.Sleep(2000);
                EventLog.PrintScreen("PlugandPlay_DeleteProjectTest_CtoG_ConstWhiteList_Page1");
                api.ByXpath("//a[contains(text(),'2')]").Click();   // page 2
                Thread.Sleep(2000);
                EventLog.PrintScreen("PlugandPlay_DeleteProjectTest_CtoG_ConstWhiteList_Page2");
            }

            // System Point
            {
                EventLog.AddLog("<GroundPC> System Point setting");
                api.ById("tagTypes").SelectTxt("System Point").Exe();
                //api.ById("SubContent").Click();   // page1
                Thread.Sleep(2000);
                EventLog.PrintScreen("PlugandPlay_DeleteProjectTest_CtoG_SystemWhiteList");
            }
            ////////////////////////////////// View Cloud White list Setting //////////////////////////////////
            PrintStep(api, "ViewandSaveGroundWhitelistSetting");

            api.Quit();
            PrintStep(api, "<GroundPC> Quit browser");
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
            long lErrorCode = 0;
            EventLog.AddLog("===PlugandPlay_DeleteProjectTest_CtoG start===");
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address(Ground PC)= " + WebAccessIP.Text);
            EventLog.AddLog("WebAccess IP address(Cloud PC)= " + WebAccessIP2.Text);
            lErrorCode = Form1_Load(ProjectName.Text, ProjectName2.Text, WebAccessIP.Text, WebAccessIP2.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===PlugandPlay_DeleteProjectTest_CtoG end===");
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
