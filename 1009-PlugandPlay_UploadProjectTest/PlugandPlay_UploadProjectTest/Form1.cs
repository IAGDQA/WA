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

namespace PlugandPlay_UploadProjectTest
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
            long lErrorCode = (long)ErrorCode.SUCCESS;
            EventLog.AddLog("===PlugandPlay_UploadProjectTest start (by iATester)===");
            if (System.IO.File.Exists(sIniFilePath))    // 再load一次
            {
                EventLog.AddLog(sIniFilePath + " file exist, load initial setting");
                InitialRequiredInfo(sIniFilePath);
            }
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(ProjectName.Text, ProjectName2.Text, WebAccessIP.Text, WebAccessIP2.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===PlugandPlay_UploadProjectTest end (by iATester)===");

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

        long Form1_Load(string sProjectName, string sProjectName2, string sWebAccessIP, string sWebAccessIP2, string sTestLogFolder, string sBrowser)
        {
            baseUrl = "http://" + sWebAccessIP;
            baseUrl2 = "http://" + sWebAccessIP2;

            // Step1: Set cloud PC mqtt broker account and password
            CloudBrokerSetting(sBrowser, sProjectName2, sWebAccessIP2, sTestLogFolder);

            // Step2: Set ground PC plug and play info and upload
            GroundPlugandPlaySetting(sBrowser, sProjectName, sProjectName2, sWebAccessIP, sWebAccessIP2, sTestLogFolder);

            EventLog.AddLog("wait 30s for cloud PC processing uploaded project...");
            Thread.Sleep(30000);    // wait for cloud processing uploaded project

            // Step3: View and save cloud PC if the project upload success
            ViewandSaveCloudProject(sBrowser, sProjectName2, sWebAccessIP2, sTestLogFolder);

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

        private void CloudBrokerSetting(string sBrowser, string sProjectName, string sWebAccessIP, string sTestLogFolder)
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
            EventLog.AddLog("<CloudPC> Set MQTT broker Account");
            EventLog.AddLog("<CloudPC> UserName=admin Password=12345");
            api2.ByXpath("//a[contains(@href, '/broadWeb/mqtt/mqttBroker.asp')]").Click();
            api2.ByName("UserName").Clear();
            api2.ByName("UserName").Enter("admin").Exe();
            api2.ByName("Password").Clear();
            api2.ByName("Password").Enter("12345").Submit().Exe();
            Thread.Sleep(500);
            PrintStep(api2, "<CloudPC> SetMQTTbrokerAccount");
            api2.Quit();
            PrintStep(api2, "<CloudPC> Quit browser");
        }

        private void GroundPlugandPlaySetting(string sBrowser, string sProjectName, string sProjectName2, string sWebAccessIP, string sWebAccessIP2, string sTestLogFolder)
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
            PrintStep(api, "<GroudPC> Login WebAccess");

            // Configure project by project name
            api.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
            PrintStep(api, "<GroundPC> Configure project");
            
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/node/nodePg.asp?') and contains(@href, 'action=node_property')]").Click();
            
            EventLog.AddLog("<GroundPC> Cloud Connection Settings");
            api.ByName("CLOUD_ENABLE").Click();
            Thread.Sleep(500);
            api.ByName("CLOUD_PROJNAME").Clear();
            api.ByName("CLOUD_PROJNAME").Enter(sProjectName2).Exe();
            api.ByName("CLOUD_SCADANAME").Clear();
            api.ByName("CLOUD_SCADANAME").Enter("CTestSCADA").Exe();
            api.ByName("DEFAULT_BUTTON").Click();
            api.ByName("CLOUD_IP").Clear();
            api.ByName("CLOUD_IP").Enter(sWebAccessIP2).Exe();
            api.ByName("CLOUD_USERNAME").Clear();
            api.ByName("CLOUD_USERNAME").Enter("admin").Exe();
            api.ByName("CLOUD_PASSWORD").Clear();
            api.ByName("CLOUD_PASSWORD").Enter("12345").Submit().Exe();
            PrintStep(api, "Cloud Connection Settings");
            
            EventLog.AddLog("<GroundPC> Cloud White list setting");
            api.SwitchToCurWindow(0);
            api.SwitchToFrame("rightFrame", 0);
            api.ByXpath("//a[contains(@href, '/broadWeb/WaCloudWhitelist/CloudWhitelist.asp?')]").Click();

            ////////////////////////////////// Cloud White list Setting //////////////////////////////////
            {   // AI/AO/DI/DO
                EventLog.AddLog("<GroundPC> Modbus tag setting");

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
                api.ByCss("img").Click();   // page1
                Thread.Sleep(2000);
                api.ByName("SetConfigAll").Click();
                api.ByName("SetDataLogAll").Click();
                api.ByName("SetDeadBandValue").Enter("0").Exe();
                SendKeys.SendWait("{ENTER}");
                Thread.Sleep(1000);
                /*
                api.ByName("SetDeadBand").Clear();
                api.ByName("SetDeadBand").Enter("0").Exe();
                
                for (int i = 2; i <= 500; i++)
                {
                    api.ByXpath("(//input[@name='SetDeadBand'])[" + i + "]").Clear();
                    api.ByXpath("(//input[@name='SetDeadBand'])[" + i + "]").Enter("0").Exe();
                }
                */
                //api.ByXpath("//input[@value='Save']").Click();
                switch (slanguage)
                {
                    case "ENG":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "CHT":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "CHS":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "JPN":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "KRN":
                        api.ByXpath("//input[@value='저장']").Click();
                        break;
                    case "FRN":
                        api.ByXpath("//input[@value='Enregistrer']").Click();
                        break;

                    default:
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                }

                Thread.Sleep(500);
                api.ByXpath("//input[@value='Ok']").Click();
                Thread.Sleep(100);

                api.ByXpath("//a[contains(text(),'2')]").Click();   // page 2
                Thread.Sleep(1000);
                api.ByName("SetConfigAll").Click();
                api.ByName("SetDataLogAll").Click();
                switch (slanguage)
                {
                    case "ENG":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "CHT":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "CHS":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "JPN":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "KRN":
                        api.ByXpath("//input[@value='저장']").Click();
                        break;
                    case "FRN":
                        api.ByXpath("//input[@value='Enregistrer']").Click();
                        break;

                    default:
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                }
                Thread.Sleep(500);
                api.ByXpath("//input[@value='Ok']").Click();
                Thread.Sleep(100);
            }
            Thread.Sleep(1000);
            // Port4(opc)
            {
                EventLog.AddLog("<GroundPC> Port4(opc) setting");
                //api.ById("tagTypes").SelectTxt("Port4(opc)").Exe();
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
                api.ByCss("img").Click();   // page1
                Thread.Sleep(1000);
                api.ByName("SetConfigAll").Click();
                api.ByName("SetDataLogAll").Click();
                api.ByName("SetDeadBandValue").Enter("0").Exe();
                SendKeys.SendWait("{ENTER}");
                Thread.Sleep(1000);
                /*
                api.ByName("SetDeadBand").Clear();
                api.ByName("SetDeadBand").Enter("0").Exe();

                for (int i = 2; i <= 250; i++)
                {
                    api.ByXpath("(//input[@name='SetDeadBand'])[" + i + "]").Clear();
                    api.ByXpath("(//input[@name='SetDeadBand'])[" + i + "]").Enter("0").Exe();
                }
                */
                switch (slanguage)
                {
                    case "ENG":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "CHT":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "CHS":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "JPN":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "KRN":
                        api.ByXpath("//input[@value='저장']").Click();
                        break;
                    case "FRN":
                        api.ByXpath("//input[@value='Enregistrer']").Click();
                        break;

                    default:
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                }
                Thread.Sleep(500);
                api.ByXpath("//input[@value='Ok']").Click();
                Thread.Sleep(100);
            }
            Thread.Sleep(1000);
            // Port5(tcpip)
            {
                EventLog.AddLog("<GroundPC> Port5(tcpip) setting");
                //api.ById("tagTypes").SelectTxt("Port5(tcpip)").Exe();
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
                api.ByCss("img").Click();   // page1
                Thread.Sleep(2000);
                api.ByName("SetConfigAll").Click();
                api.ByName("SetDataLogAll").Click();
                api.ByName("SetDeadBandValue").Enter("0").Exe();
                SendKeys.SendWait("{ENTER}");
                Thread.Sleep(1000);
                /*
                api.ByName("SetDeadBand").Clear();
                api.ByName("SetDeadBand").Enter("0").Exe();

                for (int i = 2; i <= 250; i++)
                {
                    api.ByXpath("(//input[@name='SetDeadBand'])[" + i + "]").Clear();
                    api.ByXpath("(//input[@name='SetDeadBand'])[" + i + "]").Enter("0").Exe();
                }
                */
                switch (slanguage)
                {
                    case "ENG":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "CHT":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "CHS":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "JPN":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "KRN":
                        api.ByXpath("//input[@value='저장']").Click();
                        break;
                    case "FRN":
                        api.ByXpath("//input[@value='Enregistrer']").Click();
                        break;

                    default:
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                }
                Thread.Sleep(500);
                api.ByXpath("//input[@value='Ok']").Click();
                Thread.Sleep(100);
            }
            Thread.Sleep(1000);
            // Acc Point
            {
                EventLog.AddLog("<GroundPC> Acc Point setting");
                //api.ById("tagTypes").SelectTxt("Acc Point").Exe();
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
                //api.ByCss("img").Click();   // page1
                Thread.Sleep(1000);
                api.ByName("SetConfigAll").Click();
                api.ByName("SetDataLogAll").Click();
                api.ByName("SetDeadBandValue").Enter("0").Exe();
                SendKeys.SendWait("{ENTER}");
                Thread.Sleep(1000);
                /*
                api.ByName("SetDeadBand").Clear();
                api.ByName("SetDeadBand").Enter("0").Exe();

                for (int i = 2; i <= 250; i++)
                {
                    api.ByXpath("(//input[@name='SetDeadBand'])[" + i + "]").Clear();
                    api.ByXpath("(//input[@name='SetDeadBand'])[" + i + "]").Enter("0").Exe();
                }
                */
                switch (slanguage)
                {
                    case "ENG":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "CHT":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "CHS":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "JPN":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "KRN":
                        api.ByXpath("//input[@value='저장']").Click();
                        break;
                    case "FRN":
                        api.ByXpath("//input[@value='Enregistrer']").Click();
                        break;

                    default:
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                }
                Thread.Sleep(500);
                api.ByXpath("//input[@value='Ok']").Click();
                Thread.Sleep(100);
            }
            Thread.Sleep(1000);
            // Calc Point
            {
                EventLog.AddLog("<GroundPC> Calc Point setting");
                //api.ById("tagTypes").SelectTxt("Calc Point").Exe();
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
                        api.ById("tagTypes").SelectTxt("산출 포인트").Exe(); // 翻譯可能有問題 與acc一樣
                        break;
                    case "FRN":
                        api.ById("tagTypes").SelectTxt("Point calc.").Exe();
                        break;

                    default:
                        api.ById("tagTypes").SelectTxt("Calc Point").Exe();
                        break;
                }
                //api.ByCss("img").Click();   // page1
                Thread.Sleep(1000);
                api.ByName("SetConfigAll").Click();
                api.ByName("SetDataLogAll").Click();
                api.ByName("SetDeadBandValue").Enter("0").Exe();
                SendKeys.SendWait("{ENTER}");
                Thread.Sleep(1000);

                switch (slanguage)
                {
                    case "ENG":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "CHT":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "CHS":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "JPN":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "KRN":
                        api.ByXpath("//input[@value='저장']").Click();
                        break;
                    case "FRN":
                        api.ByXpath("//input[@value='Enregistrer']").Click();
                        break;

                    default:
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                }
                Thread.Sleep(500);
                api.ByXpath("//input[@value='Ok']").Click();
                Thread.Sleep(100);
            }
            Thread.Sleep(1000);
            // Const Point
            {
                EventLog.AddLog("<GroundPC> Const Point setting");
                //api.ById("tagTypes").SelectTxt("Const Point").Exe();
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
                //api.ByCss("img").Click();   // page1
                Thread.Sleep(1000);
                api.ByName("SetConfigAll").Click();
                api.ByName("SetDataLogAll").Click();
                api.ByName("SetDeadBandValue").Enter("0").Exe();
                SendKeys.SendWait("{ENTER}");
                Thread.Sleep(1000);
                /*
                api.ByName("SetDeadBand").Clear();
                api.ByName("SetDeadBand").Enter("0").Exe();

                for (int i = 2; i <= 250; i++)
                {
                    api.ByXpath("(//input[@name='SetDeadBand'])[" + i + "]").Clear();
                    api.ByXpath("(//input[@name='SetDeadBand'])[" + i + "]").Enter("0").Exe();
                }
                */
                switch (slanguage)
                {
                    case "ENG":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "CHT":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "CHS":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "JPN":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "KRN":
                        api.ByXpath("//input[@value='저장']").Click();
                        break;
                    case "FRN":
                        api.ByXpath("//input[@value='Enregistrer']").Click();
                        break;

                    default:
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                }
                Thread.Sleep(500);
                api.ByXpath("//input[@value='Ok']").Click();
                Thread.Sleep(100);

                api.ByXpath("//a[contains(text(),'2')]").Click();   // page 2
                Thread.Sleep(1000);
                api.ByName("SetConfigAll").Click();
                api.ByName("SetDataLogAll").Click();
                switch (slanguage)
                {
                    case "ENG":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "CHT":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "CHS":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "JPN":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "KRN":
                        api.ByXpath("//input[@value='저장']").Click();
                        break;
                    case "FRN":
                        api.ByXpath("//input[@value='Enregistrer']").Click();
                        break;

                    default:
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                }
                Thread.Sleep(500);
                api.ByXpath("//input[@value='Ok']").Click();
                Thread.Sleep(100);
            }
            Thread.Sleep(1000);
            // System Point
            {
                EventLog.AddLog("<GroundPC> System Point setting");
                //api.ById("tagTypes").SelectTxt("System Point").Exe();
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
                //api.ByCss("img").Click();   // page1
                Thread.Sleep(1000);
                api.ByName("SetConfigAll").Click();
                api.ByName("SetDataLogAll").Click();
                api.ByName("SetDeadBandValue").Enter("0").Exe();
                SendKeys.SendWait("{ENTER}");
                Thread.Sleep(1000);
                /*
                api.ByName("SetDeadBand").Clear();
                api.ByName("SetDeadBand").Enter("0").Exe();

                for (int i = 2; i <= 250; i++)
                {
                    api.ByXpath("(//input[@name='SetDeadBand'])[" + i + "]").Clear();
                    api.ByXpath("(//input[@name='SetDeadBand'])[" + i + "]").Enter("0").Exe();
                }
                */
                switch (slanguage)
                {
                    case "ENG":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "CHT":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "CHS":
                        api.ByXpath("//input[@value='保存']").Click();
                        break;
                    case "JPN":
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                    case "KRN":
                        api.ByXpath("//input[@value='저장']").Click();
                        break;
                    case "FRN":
                        api.ByXpath("//input[@value='Enregistrer']").Click();
                        break;

                    default:
                        api.ByXpath("//input[@value='Save']").Click();
                        break;
                }
                Thread.Sleep(500);
                api.ByXpath("//input[@value='Ok']").Click();
                Thread.Sleep(100);
            }
            ////////////////////////////////// Cloud White list Setting //////////////////////////////////
            PrintStep(api, "CloudWhitelistSetting");
            ReturnSCADAPage(api);

            EventLog.AddLog("<GroundPC> Download...");
            StartDownload(api, sTestLogFolder);

            api.Quit();
            PrintStep(api, "Quit browser");
        }

        private void ViewandSaveCloudProject(string sBrowser, string sProjectName, string sWebAccessIP, string sTestLogFolder)
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
            Thread.Sleep(2000);
            PrintScreen("PlugandPlay_UploadProjectTest_Project Manager Page", sTestLogFolder);

            // Configure project by project name
            EventLog.AddLog("<CloudPC> Capture the configure project page");
            api2.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
            PrintStep(api2, "<CloudPC> Configure project");
            Thread.Sleep(5000);
            PrintScreen("PlugandPlay_UploadProjectTest_Configure project Page", sTestLogFolder);

            api2.Quit();
            PrintStep(api2, "<CloudPC> Quit browser");
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
            Application.DoEvents();
        }

        private void Start_Click(object sender, EventArgs e)
        {
            long lErrorCode = (long)ErrorCode.SUCCESS;
            EventLog.AddLog("===PlugandPlay_UploadProjectTest start===");
            CheckifIniFileChange();
            EventLog.AddLog("Project= " + ProjectName.Text);
            EventLog.AddLog("WebAccess IP address(Ground PC)= " + WebAccessIP.Text);
            EventLog.AddLog("WebAccess IP address(Cloud PC)= " + WebAccessIP2.Text);
            lErrorCode = Form1_Load(ProjectName.Text, ProjectName2.Text, WebAccessIP.Text, WebAccessIP2.Text, TestLogFolder.Text, Browser.Text);
            EventLog.AddLog("===PlugandPlay_UploadProjectTest end===");
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