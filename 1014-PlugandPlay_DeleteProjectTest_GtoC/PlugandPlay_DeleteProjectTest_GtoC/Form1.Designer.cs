namespace PlugandPlay_DeleteProjectTest_GtoC
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Start = new System.Windows.Forms.Button();
            this.Browser = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.ProjectName = new System.Windows.Forms.TextBox();
            this.WebAccessIP = new System.Windows.Forms.TextBox();
            this.TestLogFolder = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column_TestItem = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_BrowserAction = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_Result = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_ErrorCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column1_ExeTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.WebAccessIP2 = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.ProjectName2 = new System.Windows.Forms.TextBox();
            this.Result = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // Start
            // 
            this.Start.Location = new System.Drawing.Point(484, 36);
            this.Start.Name = "Start";
            this.Start.Size = new System.Drawing.Size(158, 130);
            this.Start.TabIndex = 0;
            this.Start.Text = "Start";
            this.Start.UseVisualStyleBackColor = true;
            this.Start.Click += new System.EventHandler(this.Start_Click);
            // 
            // Browser
            // 
            this.Browser.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Browser.FormattingEnabled = true;
            this.Browser.Items.AddRange(new object[] {
            "Internet Explorer",
            "Mozilla FireFox",
            "Chrome",
            "Safari"});
            this.Browser.Location = new System.Drawing.Point(135, 168);
            this.Browser.Name = "Browser";
            this.Browser.Size = new System.Drawing.Size(121, 20);
            this.Browser.TabIndex = 1;
            this.Browser.Tag = "";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(26, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 12);
            this.label1.TabIndex = 2;
            this.label1.Text = "Project Name";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(26, 66);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 24);
            this.label2.TabIndex = 3;
            this.label2.Text = "WebAccess IP\r\n (Ground PC)";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(26, 138);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(79, 12);
            this.label3.TabIndex = 4;
            this.label3.Text = "Test Log Folder";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(26, 174);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(44, 12);
            this.label4.TabIndex = 5;
            this.label4.Text = "Browser";
            // 
            // ProjectName
            // 
            this.ProjectName.Location = new System.Drawing.Point(135, 36);
            this.ProjectName.Name = "ProjectName";
            this.ProjectName.Size = new System.Drawing.Size(100, 22);
            this.ProjectName.TabIndex = 6;
            this.ProjectName.Text = "TestProject";
            // 
            // WebAccessIP
            // 
            this.WebAccessIP.Location = new System.Drawing.Point(135, 69);
            this.WebAccessIP.Name = "WebAccessIP";
            this.WebAccessIP.Size = new System.Drawing.Size(290, 22);
            this.WebAccessIP.TabIndex = 7;
            this.WebAccessIP.Text = "172.18.3.69";
            // 
            // TestLogFolder
            // 
            this.TestLogFolder.Location = new System.Drawing.Point(135, 133);
            this.TestLogFolder.Name = "TestLogFolder";
            this.TestLogFolder.Size = new System.Drawing.Size(290, 22);
            this.TestLogFolder.TabIndex = 8;
            this.TestLogFolder.Text = "C:\\WALogData";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column_TestItem,
            this.Column_BrowserAction,
            this.Column_Result,
            this.Column_ErrorCode,
            this.Column1_ExeTime});
            this.dataGridView1.Location = new System.Drawing.Point(28, 200);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(643, 194);
            this.dataGridView1.TabIndex = 9;
            // 
            // Column_TestItem
            // 
            this.Column_TestItem.FillWeight = 150F;
            this.Column_TestItem.HeaderText = "Test Item";
            this.Column_TestItem.Name = "Column_TestItem";
            this.Column_TestItem.Width = 150;
            // 
            // Column_BrowserAction
            // 
            this.Column_BrowserAction.FillWeight = 200F;
            this.Column_BrowserAction.HeaderText = "Browser Action";
            this.Column_BrowserAction.Name = "Column_BrowserAction";
            this.Column_BrowserAction.Width = 200;
            // 
            // Column_Result
            // 
            this.Column_Result.FillWeight = 50F;
            this.Column_Result.HeaderText = "Result";
            this.Column_Result.Name = "Column_Result";
            this.Column_Result.Width = 50;
            // 
            // Column_ErrorCode
            // 
            this.Column_ErrorCode.HeaderText = "Error Code";
            this.Column_ErrorCode.Name = "Column_ErrorCode";
            // 
            // Column1_ExeTime
            // 
            this.Column1_ExeTime.HeaderText = "Exe Time (ms)";
            this.Column1_ExeTime.Name = "Column1_ExeTime";
            // 
            // WebAccessIP2
            // 
            this.WebAccessIP2.Location = new System.Drawing.Point(135, 102);
            this.WebAccessIP2.Name = "WebAccessIP2";
            this.WebAccessIP2.Size = new System.Drawing.Size(290, 22);
            this.WebAccessIP2.TabIndex = 10;
            this.WebAccessIP2.Text = "172.18.3.65";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(26, 101);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(71, 24);
            this.label5.TabIndex = 11;
            this.label5.Text = "WebAccess IP\r\n  (Cloud PC)";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(247, 36);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(67, 24);
            this.label6.TabIndex = 12;
            this.label6.Text = "Project Name\r\n    (Cloud)";
            // 
            // ProjectName2
            // 
            this.ProjectName2.Location = new System.Drawing.Point(325, 36);
            this.ProjectName2.Name = "ProjectName2";
            this.ProjectName2.Size = new System.Drawing.Size(100, 22);
            this.ProjectName2.TabIndex = 13;
            this.ProjectName2.Text = "CTestProject";
            // 
            // Result
            // 
            this.Result.AutoSize = true;
            this.Result.Location = new System.Drawing.Point(323, 174);
            this.Result.Name = "Result";
            this.Result.Size = new System.Drawing.Size(35, 12);
            this.Result.TabIndex = 14;
            this.Result.Text = "Ready";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(698, 412);
            this.Controls.Add(this.Result);
            this.Controls.Add(this.ProjectName2);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.WebAccessIP2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.TestLogFolder);
            this.Controls.Add(this.WebAccessIP);
            this.Controls.Add(this.ProjectName);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Browser);
            this.Controls.Add(this.Start);
            this.Name = "Form1";
            this.Text = "Advantech WebAccess Auto Test (PlugandPlay_DeleteProjectTest <Ground to Cloud>)";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Start;
        private System.Windows.Forms.ComboBox Browser;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox ProjectName;
        private System.Windows.Forms.TextBox WebAccessIP;
        private System.Windows.Forms.TextBox TestLogFolder;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_TestItem;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_BrowserAction;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_Result;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_ErrorCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1_ExeTime;
        private System.Windows.Forms.TextBox WebAccessIP2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox ProjectName2;
        private System.Windows.Forms.Label Result;
    }
}

