namespace CreateRecipe
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
            this.TestLogFolder = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.Value = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.Recipe_Name = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.Unit_Name = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.Recipe_File_Name = new System.Windows.Forms.TextBox();
            this.Column_TestItem = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_BrowserAction = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column_Result = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label4 = new System.Windows.Forms.Label();
            this.Column_ExeTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label2 = new System.Windows.Forms.Label();
            this.Result = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.Column_ErrorCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Browser = new System.Windows.Forms.ComboBox();
            this.WebAccessIP = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // TestLogFolder
            // 
            this.TestLogFolder.Location = new System.Drawing.Point(177, 83);
            this.TestLogFolder.Name = "TestLogFolder";
            this.TestLogFolder.Size = new System.Drawing.Size(246, 22);
            this.TestLogFolder.TabIndex = 57;
            this.TestLogFolder.Text = "C:\\WALogData";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(88, 83);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(79, 12);
            this.label1.TabIndex = 56;
            this.label1.Text = "Test Log Folder";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(96, 197);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(32, 12);
            this.label6.TabIndex = 55;
            this.label6.Text = "Value";
            // 
            // Value
            // 
            this.Value.Location = new System.Drawing.Point(177, 194);
            this.Value.Name = "Value";
            this.Value.Size = new System.Drawing.Size(246, 22);
            this.Value.TabIndex = 54;
            this.Value.Text = "500";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(96, 169);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(67, 12);
            this.label11.TabIndex = 53;
            this.label11.Text = "Recipe Name";
            // 
            // Recipe_Name
            // 
            this.Recipe_Name.Location = new System.Drawing.Point(177, 166);
            this.Recipe_Name.Name = "Recipe_Name";
            this.Recipe_Name.Size = new System.Drawing.Size(246, 22);
            this.Recipe_Name.TabIndex = 52;
            this.Recipe_Name.Text = "Recipe1";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(112, 142);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(55, 12);
            this.label5.TabIndex = 51;
            this.label5.Text = "Unit Name";
            // 
            // Unit_Name
            // 
            this.Unit_Name.Location = new System.Drawing.Point(177, 139);
            this.Unit_Name.Name = "Unit_Name";
            this.Unit_Name.Size = new System.Drawing.Size(246, 22);
            this.Unit_Name.TabIndex = 50;
            this.Unit_Name.Text = "Unit1";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(80, 114);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(87, 12);
            this.label3.TabIndex = 49;
            this.label3.Text = "Recipe File Name";
            // 
            // Recipe_File_Name
            // 
            this.Recipe_File_Name.Location = new System.Drawing.Point(177, 111);
            this.Recipe_File_Name.Name = "Recipe_File_Name";
            this.Recipe_File_Name.Size = new System.Drawing.Size(246, 22);
            this.Recipe_File_Name.TabIndex = 48;
            this.Recipe_File_Name.Text = "test";
            // 
            // Column_TestItem
            // 
            this.Column_TestItem.HeaderText = "Test Item";
            this.Column_TestItem.Name = "Column_TestItem";
            // 
            // Column_BrowserAction
            // 
            this.Column_BrowserAction.HeaderText = "Browser Action";
            this.Column_BrowserAction.Name = "Column_BrowserAction";
            // 
            // Column_Result
            // 
            this.Column_Result.HeaderText = "Result";
            this.Column_Result.Name = "Column_Result";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(123, 55);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(44, 12);
            this.label4.TabIndex = 47;
            this.label4.Text = "Browser";
            // 
            // Column_ExeTime
            // 
            this.Column_ExeTime.HeaderText = "Exe Time (ms)";
            this.Column_ExeTime.Name = "Column_ExeTime";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(96, 27);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 12);
            this.label2.TabIndex = 46;
            this.label2.Text = "WebAccess IP";
            // 
            // Result
            // 
            this.Result.AutoSize = true;
            this.Result.Location = new System.Drawing.Point(510, 166);
            this.Result.Name = "Result";
            this.Result.Size = new System.Drawing.Size(35, 12);
            this.Result.TabIndex = 45;
            this.Result.Text = "Ready";
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Column_TestItem,
            this.Column_BrowserAction,
            this.Column_Result,
            this.Column_ErrorCode,
            this.Column_ExeTime});
            this.dataGridView1.Location = new System.Drawing.Point(70, 228);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(638, 195);
            this.dataGridView1.TabIndex = 44;
            // 
            // Column_ErrorCode
            // 
            this.Column_ErrorCode.HeaderText = "Error Code";
            this.Column_ErrorCode.Name = "Column_ErrorCode";
            // 
            // Browser
            // 
            this.Browser.FormattingEnabled = true;
            this.Browser.Items.AddRange(new object[] {
            "Internet Explorer",
            "Mozilla FireFox"});
            this.Browser.Location = new System.Drawing.Point(177, 52);
            this.Browser.Name = "Browser";
            this.Browser.Size = new System.Drawing.Size(121, 20);
            this.Browser.TabIndex = 43;
            // 
            // WebAccessIP
            // 
            this.WebAccessIP.Location = new System.Drawing.Point(177, 24);
            this.WebAccessIP.Name = "WebAccessIP";
            this.WebAccessIP.Size = new System.Drawing.Size(246, 22);
            this.WebAccessIP.TabIndex = 42;
            this.WebAccessIP.Text = "172.16.12.138";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(493, 50);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(109, 83);
            this.button1.TabIndex = 58;
            this.button1.Text = "Start";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(778, 446);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.TestLogFolder);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.Value);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.Recipe_Name);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.Unit_Name);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.Recipe_File_Name);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.Result);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.Browser);
            this.Controls.Add(this.WebAccessIP);
            this.Name = "Form1";
            this.Text = "Advantech WebAccess Auto Test ( CreateRecipe)";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox TestLogFolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox Value;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox Recipe_Name;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox Unit_Name;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox Recipe_File_Name;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_TestItem;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_BrowserAction;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_Result;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_ExeTime;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label Result;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column_ErrorCode;
        private System.Windows.Forms.ComboBox Browser;
        private System.Windows.Forms.TextBox WebAccessIP;
        private System.Windows.Forms.Button button1;
    }
}

