using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AdvWebUIAPI;
using ThirdPartyToolControl;
using iATester;
using System.Runtime.InteropServices;

namespace CreateRecipe
{
    public partial class Form1 : Form, iATester.iCom
    {
        IAdvSeleniumAPI api;
        cThirdPartyToolControl tpc = new cThirdPartyToolControl();
        private delegate void DataGridViewCtrlAddDataRow(DataGridViewRow i_Row);
        private DataGridViewCtrlAddDataRow m_DataGridViewCtrlAddDataRow;
        internal const int Max_Rows_Val = 65535;
        string baseUrl;
        string sIniFilePath = @"C:\WebAccessAutoTestSetting.ini";

        string slanguage;
        string sProjectName;

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
            EventLog.AddLog("===Create Recipe start===");
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(WebAccessIP.Text, Browser.Text, Recipe_File_Name.Text, Unit_Name.Text, Recipe_Name.Text, Value.Text);
            EventLog.AddLog("===Create Recipe end===");

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

            //eStatus(this, new StatusEventArgs(iStatus.Completion));
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

        private void InitialRequiredInfo(string sFilePath)
        {
            StringBuilder sDefaultUserEmail = new StringBuilder(255);
            StringBuilder sDefaultProjectName1 = new StringBuilder(255);
            StringBuilder sDefaultProjectName2 = new StringBuilder(255);
            StringBuilder sDefaultIP1 = new StringBuilder(255);
            StringBuilder sDefaultIP2 = new StringBuilder(255);
            StringBuilder sDefaultUserLanguage = new StringBuilder(255);

            tpc.F_GetPrivateProfileString("UserInfo", "Email", "NA", sDefaultUserEmail, 255, sFilePath);
            tpc.F_GetPrivateProfileString("UserInfo", "Language", "NA", sDefaultUserLanguage, 255, sFilePath);
            tpc.F_GetPrivateProfileString("ProjectName", "Ground PC or Primary PC", "NA", sDefaultProjectName1, 255, sFilePath);
            tpc.F_GetPrivateProfileString("ProjectName", "Cloud PC or Backup PC", "NA", sDefaultProjectName2, 255, sFilePath);
            tpc.F_GetPrivateProfileString("IP", "Ground PC or Primary PC", "NA", sDefaultIP1, 255, sFilePath);
            tpc.F_GetPrivateProfileString("IP", "Cloud PC or Backup PC", "NA", sDefaultIP2, 255, sFilePath);

            WebAccessIP.Text = sDefaultIP1.ToString();
            slanguage = sDefaultUserLanguage.ToString();
            sProjectName = sDefaultProjectName1.ToString();
        }

        long Form1_Load(string sWebAccessIP, string sBrowser, string Recipe_File_Name, string Unit_Name, string Recipe_Name, string Value)
        {
            try
            {
                baseUrl = "http://" + sWebAccessIP;
                if (sBrowser == "Internet Explorer")
                {
                    EventLog.AddLog("Browser= Internet Explorer");
                    //driver = new FirefoxDriver();
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
                //api.ById("userField").Enter("").Submit().Exe();
                api.ByXpath("//input[@id='submit1']").Click();   //??
                PrintStep("Login WebAccess");

                //Step0: select first project and scada
                EventLog.AddLog("Create a recipe");
                api.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp') and contains(@href, 'ProjName=" + sProjectName + "')]").Click();
                //iCheckIfSCADAExis = api.ByXpath("//a[contains(@href, '/broadWeb/bwMainRight.asp') and contains(@href, 'name=TestSCADA')]").Click();
                //api.ByXpath("//a[contains(@href, '/broadWeb/bwMain.asp?pos=project&pid=2&ProjName=TestProjectGGGG')]").Click();      //??

                //Step1: select recipe and add new recipe
                api.SwitchToFrame("rightFrame", 0);
                api.ByXpath("//a[contains(@href, '/broadWeb/recipe/rpList.asp')]").Click();
                api.ByXpath("//a[contains(@href, '/broadWeb/recipe/rpPg.asp') and contains(@href, 'action=add_recipe')]").Click();

                //Step2: set recipe data
                bool bResult = CreateRecipe(Recipe_File_Name, Unit_Name, Recipe_Name, Value);
                PrintStep("create a recipe");

                api.Quit();

                if (bResult)
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
            }
            catch (Exception ex)
            {
                EventLog.AddLog(ex.ToString());
                return -1;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            long lErrorCode = (long)ErrorCode.SUCCESS;
            EventLog.AddLog("===Create Recipe start===");
            EventLog.AddLog("WebAccess IP address= " + WebAccessIP.Text);
            lErrorCode = Form1_Load(WebAccessIP.Text, Browser.Text, Recipe_File_Name.Text, Unit_Name.Text, Recipe_Name.Text, Value.Text);
            EventLog.AddLog("===Create Recipe end===");
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
        }

        bool CreateRecipe(string Recipe_File_Name, string Unit_Name, string Recipe_Name, string Value)
        {
            try
            {
                api.ByName("FileName").Clear();
                api.ByName("FileName").Enter(Recipe_File_Name).Exe();
                api.ByName("UnitName").Clear();
                api.ByName("UnitName").Enter(Unit_Name).Exe();
                api.ByName("RecipeName").Clear();
                api.ByName("RecipeName").Enter(Recipe_Name).Exe();
                api.ByName("ItemName_1").Clear();
                api.ByName("ItemName_1").Enter("ItemName_1").Exe();
                api.ByName("ItemName_2").Clear();
                api.ByName("ItemName_2").Enter("ItemName_2").Exe();
                api.ByName("TagName001").Clear();
                api.ByName("TagName001").Enter("ConAna_0249").Exe();
                api.ByName("TagName002").Clear();
                api.ByName("TagName002").Enter("ConAna_0250").Exe();
                api.ByName("PreValue_1").Clear();
                api.ByName("PreValue_1").Enter(Value).Exe();
                api.ByName("PreValue_2").Clear();
                api.ByName("PreValue_2").Enter(Value).Exe();
                int iSubmitResult = api.ByName("PreValue_3").Enter("").Submit().Exe();
                if (iSubmitResult == 0)
                {
                    EventLog.AddLog("Create success!!");
                    return true;
                }
                else
                {
                    EventLog.AddLog("Create fail!!");
                    return false;
                }
            }
            catch (Exception ex)
            {
                EventLog.AddLog(ex.ToString());
                return false;
            }
        }
    }
}
