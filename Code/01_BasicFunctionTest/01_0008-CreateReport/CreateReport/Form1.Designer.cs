namespace CreateReport
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
            this.comboBox_Browser = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox_Primary_project = new System.Windows.Forms.TextBox();
            this.textBox_Primary_IP = new System.Windows.Forms.TextBox();
            this.Result = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.textbox_UserEmail = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textBox_Secondary_IP = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.comboBox_Language = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox_Secondary_project = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // Start
            // 
            this.Start.Location = new System.Drawing.Point(473, 36);
            this.Start.Name = "Start";
            this.Start.Size = new System.Drawing.Size(145, 145);
            this.Start.TabIndex = 8;
            this.Start.Text = "Start";
            this.Start.UseVisualStyleBackColor = true;
            this.Start.Click += new System.EventHandler(this.Start_Click);
            // 
            // comboBox_Browser
            // 
            this.comboBox_Browser.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Browser.FormattingEnabled = true;
            this.comboBox_Browser.Items.AddRange(new object[] {
            "Internet Explorer",
            "Mozilla FireFox",
            "Chrome",
            "Safari",
            "Edge"});
            this.comboBox_Browser.Location = new System.Drawing.Point(167, 180);
            this.comboBox_Browser.Name = "comboBox_Browser";
            this.comboBox_Browser.Size = new System.Drawing.Size(121, 20);
            this.comboBox_Browser.TabIndex = 6;
            this.comboBox_Browser.Tag = "";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(26, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(105, 12);
            this.label1.TabIndex = 9;
            this.label1.Text = "Primary project name";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(26, 69);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 12);
            this.label2.TabIndex = 10;
            this.label2.Text = "Primary IP";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(26, 183);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(44, 12);
            this.label4.TabIndex = 14;
            this.label4.Text = "Browser";
            // 
            // textBox_Primary_project
            // 
            this.textBox_Primary_project.Location = new System.Drawing.Point(167, 36);
            this.textBox_Primary_project.MaxLength = 15;
            this.textBox_Primary_project.Name = "textBox_Primary_project";
            this.textBox_Primary_project.Size = new System.Drawing.Size(258, 22);
            this.textBox_Primary_project.TabIndex = 1;
            this.textBox_Primary_project.Text = "TestProjectGo01";
            // 
            // textBox_Primary_IP
            // 
            this.textBox_Primary_IP.Location = new System.Drawing.Point(167, 66);
            this.textBox_Primary_IP.MaxLength = 15;
            this.textBox_Primary_IP.Name = "textBox_Primary_IP";
            this.textBox_Primary_IP.Size = new System.Drawing.Size(258, 22);
            this.textBox_Primary_IP.TabIndex = 2;
            this.textBox_Primary_IP.Text = "172.xx.xx.xx";
            // 
            // Result
            // 
            this.Result.AutoSize = true;
            this.Result.Location = new System.Drawing.Point(343, 199);
            this.Result.Name = "Result";
            this.Result.Size = new System.Drawing.Size(35, 12);
            this.Result.TabIndex = 16;
            this.Result.Text = "Ready";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(26, 154);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(56, 12);
            this.label5.TabIndex = 13;
            this.label5.Text = "User Email";
            // 
            // textbox_UserEmail
            // 
            this.textbox_UserEmail.Location = new System.Drawing.Point(167, 151);
            this.textbox_UserEmail.Name = "textbox_UserEmail";
            this.textbox_UserEmail.Size = new System.Drawing.Size(258, 22);
            this.textbox_UserEmail.TabIndex = 5;
            this.textbox_UserEmail.Text = "xxx@advantech.com.tw";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(26, 126);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(67, 12);
            this.label6.TabIndex = 12;
            this.label6.Text = "Secondary IP";
            // 
            // textBox_Secondary_IP
            // 
            this.textBox_Secondary_IP.Location = new System.Drawing.Point(167, 123);
            this.textBox_Secondary_IP.MaxLength = 15;
            this.textBox_Secondary_IP.Name = "textBox_Secondary_IP";
            this.textBox_Secondary_IP.Size = new System.Drawing.Size(258, 22);
            this.textBox_Secondary_IP.TabIndex = 4;
            this.textBox_Secondary_IP.Text = "172.xx.xx.xx";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(26, 209);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(51, 12);
            this.label8.TabIndex = 15;
            this.label8.Text = "Language";
            // 
            // comboBox_Language
            // 
            this.comboBox_Language.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_Language.FormattingEnabled = true;
            this.comboBox_Language.Items.AddRange(new object[] {
            "ENG",
            "CHT",
            "CHS",
            "JPN",
            "KRN",
            "FRN"});
            this.comboBox_Language.Location = new System.Drawing.Point(167, 206);
            this.comboBox_Language.Name = "comboBox_Language";
            this.comboBox_Language.Size = new System.Drawing.Size(121, 20);
            this.comboBox_Language.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(26, 97);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(117, 12);
            this.label3.TabIndex = 11;
            this.label3.Text = "Secondary project name";
            // 
            // textBox_Secondary_project
            // 
            this.textBox_Secondary_project.Location = new System.Drawing.Point(167, 95);
            this.textBox_Secondary_project.MaxLength = 15;
            this.textBox_Secondary_project.Name = "textBox_Secondary_project";
            this.textBox_Secondary_project.Size = new System.Drawing.Size(258, 22);
            this.textBox_Secondary_project.TabIndex = 3;
            this.textBox_Secondary_project.Text = "TestProjectGo02";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(649, 247);
            this.Controls.Add(this.textBox_Secondary_project);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.comboBox_Language);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.textBox_Secondary_IP);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.textbox_UserEmail);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.Result);
            this.Controls.Add(this.textBox_Primary_IP);
            this.Controls.Add(this.textBox_Primary_project);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox_Browser);
            this.Controls.Add(this.Start);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Advantech WebAccess Auto Test";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Start;
        private System.Windows.Forms.ComboBox comboBox_Browser;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBox_Primary_project;
        private System.Windows.Forms.TextBox textBox_Primary_IP;
        private System.Windows.Forms.Label Result;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textbox_UserEmail;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textBox_Secondary_IP;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox comboBox_Language;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox_Secondary_project;
    }
}


