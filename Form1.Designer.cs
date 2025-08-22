namespace AutoPressApp
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnStart = new System.Windows.Forms.Button();
            this.lblStartInfo = new System.Windows.Forms.Label();
            this.btnWorkflowRecord = new System.Windows.Forms.Button(); // legacy (hidden)
            this.btnWorkflowRun = new System.Windows.Forms.Button();    // legacy (hidden)
            this.btnWorkflowMenu = new System.Windows.Forms.Button();
            // 移除單鍵/間隔相關控制項
            this.lblStatus = new System.Windows.Forms.Label();
            this.lblMode = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkTestMode = new System.Windows.Forms.CheckBox();
            this.cmbApplications = new System.Windows.Forms.ComboBox();
            this.lblApplication = new System.Windows.Forms.Label();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnRecord = new System.Windows.Forms.Button();
            this.btnReplay = new System.Windows.Forms.Button();
            this.btnClearRecord = new System.Windows.Forms.Button();
            this.lstRecordedKeys = new System.Windows.Forms.ListBox();
            this.lblRecordedKeys = new System.Windows.Forms.Label();
            this.chkLoop = new System.Windows.Forms.CheckBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnImport = new System.Windows.Forms.Button();
            this.cmbSpeed = new System.Windows.Forms.ComboBox();
            this.lblSpeed = new System.Windows.Forms.Label();
            this.groupBoxHelp = new System.Windows.Forms.GroupBox();
            this.rtbHelp = new System.Windows.Forms.RichTextBox();
            this.groupBoxTest = new System.Windows.Forms.GroupBox();
            this.txtTest = new System.Windows.Forms.TextBox();
            this.btnClearTest = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBoxHelp.SuspendLayout();
            this.groupBoxTest.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnStart
            // 
            this.btnStart.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(120)))), ((int)(((byte)(215)))));
            this.btnStart.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnStart.Font = new System.Drawing.Font("Microsoft JhengHei", 14F, System.Drawing.FontStyle.Bold);
            this.btnStart.ForeColor = System.Drawing.Color.White;
            this.btnStart.Location = new System.Drawing.Point(500, 150);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(150, 50);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "▶ 開始執行";
            this.btnStart.UseVisualStyleBackColor = false;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);

            // btnWorkflowRun
            //
            this.btnWorkflowRun.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.btnWorkflowRun.Location = new System.Drawing.Point(500, 210);
            this.btnWorkflowRun.Name = "btnWorkflowRun";
            this.btnWorkflowRun.Size = new System.Drawing.Size(150, 28);
            this.btnWorkflowRun.TabIndex = 26;
            this.btnWorkflowRun.Text = "執行流程 (JSON)";
            this.btnWorkflowRun.UseVisualStyleBackColor = true;
            this.btnWorkflowRun.Visible = false; // replaced by menu button

            // btnWorkflowRecord
            //
            this.btnWorkflowRecord.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.btnWorkflowRecord.Location = new System.Drawing.Point(500, 245);
            this.btnWorkflowRecord.Name = "btnWorkflowRecord";
            this.btnWorkflowRecord.Size = new System.Drawing.Size(150, 28);
            this.btnWorkflowRecord.TabIndex = 27;
            this.btnWorkflowRecord.Text = "錄製流程 (滑鼠+視窗)";
            this.btnWorkflowRecord.UseVisualStyleBackColor = true;
            this.btnWorkflowRecord.Visible = false; // replaced by menu button

            // btnWorkflowMenu
            //
            this.btnWorkflowMenu.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.btnWorkflowMenu.Location = new System.Drawing.Point(500, 210);
            this.btnWorkflowMenu.Name = "btnWorkflowMenu";
            this.btnWorkflowMenu.Size = new System.Drawing.Size(150, 63);
            this.btnWorkflowMenu.TabIndex = 28;
            this.btnWorkflowMenu.Text = "流程功能 ▼";
            this.btnWorkflowMenu.UseVisualStyleBackColor = true;
            this.btnWorkflowMenu.Click += new System.EventHandler(this.btnWorkflowMenu_Click);
            // 
            // lblStartInfo
            // 
            this.lblStartInfo.AutoSize = true;
            this.lblStartInfo.Font = new System.Drawing.Font("Microsoft JhengHei", 10F);
            this.lblStartInfo.ForeColor = System.Drawing.Color.DarkBlue;
            this.lblStartInfo.Location = new System.Drawing.Point(500, 120);
            this.lblStartInfo.Name = "lblStartInfo";
            this.lblStartInfo.Size = new System.Drawing.Size(150, 18);
            this.lblStartInfo.TabIndex = 25;
            this.lblStartInfo.Text = "執行已記錄的按鍵序列";
            // 
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.lblStatus.ForeColor = System.Drawing.Color.Blue;
            this.lblStatus.Location = new System.Drawing.Point(30, 560);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(67, 16);
            this.lblStatus.TabIndex = 5;
            this.lblStatus.Text = "狀態: 待命中";
            // 
            // lblMode
            // 
            this.lblMode.AutoSize = true;
            this.lblMode.Font = new System.Drawing.Font("Microsoft JhengHei", 10F, System.Drawing.FontStyle.Bold);
            this.lblMode.Location = new System.Drawing.Point(500, 90);
            this.lblMode.Name = "lblMode";
            this.lblMode.Padding = new System.Windows.Forms.Padding(6, 3, 6, 3);
            this.lblMode.Size = new System.Drawing.Size(86, 24);
            this.lblMode.TabIndex = 29;
            this.lblMode.Text = "狀態: 空閒";
            this.lblMode.BackColor = System.Drawing.SystemColors.ControlLight;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnRefresh);
            this.groupBox1.Controls.Add(this.lblApplication);
            this.groupBox1.Controls.Add(this.cmbApplications);
            this.groupBox1.Controls.Add(this.chkTestMode);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft JhengHei", 10F);
            this.groupBox1.Location = new System.Drawing.Point(30, 30);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(450, 160);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "基本設定 (僅視窗目標)";
            // 
            // chkTestMode
            // 
            this.chkTestMode.AutoSize = true;
            this.chkTestMode.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.chkTestMode.Location = new System.Drawing.Point(30, 60);
            this.chkTestMode.Name = "chkTestMode";
            this.chkTestMode.Size = new System.Drawing.Size(87, 20);
            this.chkTestMode.TabIndex = 10;
            this.chkTestMode.Text = "測試模式";
            this.chkTestMode.UseVisualStyleBackColor = true;
            this.chkTestMode.Visible = true;
            // 
            // cmbApplications
            // 
            this.cmbApplications.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbApplications.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.cmbApplications.FormattingEnabled = true;
            this.cmbApplications.Location = new System.Drawing.Point(100, 110);
            this.cmbApplications.Name = "cmbApplications";
            this.cmbApplications.Size = new System.Drawing.Size(250, 24);
            this.cmbApplications.TabIndex = 7;
            // 
            // lblApplication
            // 
            this.lblApplication.AutoSize = true;
            this.lblApplication.Font = new System.Drawing.Font("Microsoft JhengHei", 10F);
            this.lblApplication.Location = new System.Drawing.Point(30, 113);
            this.lblApplication.Name = "lblApplication";
            this.lblApplication.Size = new System.Drawing.Size(79, 18);
            this.lblApplication.TabIndex = 8;
            this.lblApplication.Text = "目標視窗:";
            // 
            // btnRefresh
            // 
            this.btnRefresh.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.btnRefresh.Location = new System.Drawing.Point(360, 110);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(60, 24);
            this.btnRefresh.TabIndex = 9;
            this.btnRefresh.Text = "重新整理";
            this.btnRefresh.UseVisualStyleBackColor = true;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.lblRecordedKeys);
            this.groupBox2.Controls.Add(this.lstRecordedKeys);
            this.groupBox2.Controls.Add(this.btnClearRecord);
            this.groupBox2.Controls.Add(this.btnReplay);
            this.groupBox2.Controls.Add(this.btnRecord);
            this.groupBox2.Controls.Add(this.chkLoop);
            this.groupBox2.Controls.Add(this.btnExport);
            this.groupBox2.Controls.Add(this.btnImport);
            this.groupBox2.Controls.Add(this.cmbSpeed);
            this.groupBox2.Controls.Add(this.lblSpeed);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft JhengHei", 10F);
            this.groupBox2.Location = new System.Drawing.Point(30, 210);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(450, 270);
            this.groupBox2.TabIndex = 10;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "按鍵記錄與回放";
            // 
            // btnExport
            // 
            this.btnExport.Enabled = true;
            this.btnExport.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.btnExport.Location = new System.Drawing.Point(30, 210);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(80, 25);
            this.btnExport.TabIndex = 17;
            this.btnExport.Text = "匯出";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnImport
            // 
            this.btnImport.Enabled = true;
            this.btnImport.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.btnImport.Location = new System.Drawing.Point(120, 210);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(80, 25);
            this.btnImport.TabIndex = 18;
            this.btnImport.Text = "匯入";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // cmbSpeed
            // 
            this.cmbSpeed.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSpeed.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.cmbSpeed.FormattingEnabled = true;
            this.cmbSpeed.Items.AddRange(new object[] {"0.05x","0.1x","0.2x","0.5x","1.0x","1.5x","2.0x","3.0x","4.0x","5.0x","6.0x","8.0x","10.0x"});
            this.cmbSpeed.Location = new System.Drawing.Point(330, 235);
            this.cmbSpeed.Name = "cmbSpeed";
            this.cmbSpeed.Size = new System.Drawing.Size(90, 24);
            this.cmbSpeed.TabIndex = 19;
            this.cmbSpeed.SelectedIndex = 4; // default 1.0x
            this.cmbSpeed.SelectedIndexChanged += new System.EventHandler(this.cmbSpeed_SelectedIndexChanged);
            // 
            // lblSpeed
            // 
            this.lblSpeed.AutoSize = true;
            this.lblSpeed.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.lblSpeed.Location = new System.Drawing.Point(250, 239);
            this.lblSpeed.Name = "lblSpeed";
            this.lblSpeed.Size = new System.Drawing.Size(79, 16);
            this.lblSpeed.TabIndex = 20;
            this.lblSpeed.Text = "速度倍率:";
            // 
            // groupBoxHelp
            // 
            this.groupBoxHelp.Controls.Add(this.rtbHelp);
            this.groupBoxHelp.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.groupBoxHelp.Location = new System.Drawing.Point(30, 490);
            this.groupBoxHelp.Name = "groupBoxHelp";
            this.groupBoxHelp.Size = new System.Drawing.Size(620, 160);
            this.groupBoxHelp.TabIndex = 17;
            this.groupBoxHelp.TabStop = false;
            this.groupBoxHelp.Text = "使用說明 / 熱鍵";
            // 
            // groupBoxTest
            // 
            this.groupBoxTest.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxTest.Controls.Add(this.btnClearTest);
            this.groupBoxTest.Controls.Add(this.txtTest);
            this.groupBoxTest.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.groupBoxTest.Location = new System.Drawing.Point(30, 670);
            this.groupBoxTest.Name = "groupBoxTest";
            this.groupBoxTest.Size = new System.Drawing.Size(620, 140);
            this.groupBoxTest.TabIndex = 21;
            this.groupBoxTest.TabStop = false;
            this.groupBoxTest.Text = "測試輸入 (啟用測試模式時回放寫入此框)";
            // 
            // txtTest
            // 
            this.txtTest.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtTest.Font = new System.Drawing.Font("Consolas", 10F);
            this.txtTest.Location = new System.Drawing.Point(15, 22);
            this.txtTest.Multiline = true;
            this.txtTest.Name = "txtTest";
            this.txtTest.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtTest.Size = new System.Drawing.Size(590, 80);
            this.txtTest.TabIndex = 0;
            // 
            // btnClearTest
            // 
            this.btnClearTest.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClearTest.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.btnClearTest.Location = new System.Drawing.Point(525, 108);
            this.btnClearTest.Name = "btnClearTest";
            this.btnClearTest.Size = new System.Drawing.Size(80, 25);
            this.btnClearTest.TabIndex = 1;
            this.btnClearTest.Text = "清除";
            this.btnClearTest.UseVisualStyleBackColor = true;
            this.btnClearTest.Click += new System.EventHandler(this.btnClearTest_Click);
            // rtbHelp
            // 
            this.rtbHelp.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.rtbHelp.Location = new System.Drawing.Point(15, 22);
            this.rtbHelp.Name = "rtbHelp";
            this.rtbHelp.ReadOnly = true;
            this.rtbHelp.Size = new System.Drawing.Size(590, 125);
            this.rtbHelp.TabIndex = 0;
            this.rtbHelp.Text = "";
            this.rtbHelp.BackColor = System.Drawing.SystemColors.Window;
            this.rtbHelp.DetectUrls = false;
            this.rtbHelp.ShortcutsEnabled = false;
            this.rtbHelp.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            // 
            // btnRecord
            // 
            this.btnRecord.Font = new System.Drawing.Font("Microsoft JhengHei", 10F, System.Drawing.FontStyle.Bold);
            this.btnRecord.Location = new System.Drawing.Point(30, 30);
            this.btnRecord.Name = "btnRecord";
            this.btnRecord.Size = new System.Drawing.Size(100, 35);
            this.btnRecord.TabIndex = 11;
            this.btnRecord.Text = "開始記錄";
            this.btnRecord.UseVisualStyleBackColor = true;
            this.btnRecord.Click += new System.EventHandler(this.btnRecord_Click);
            // 
            // btnReplay
            // 
            this.btnReplay.Enabled = false;
            this.btnReplay.Font = new System.Drawing.Font("Microsoft JhengHei", 10F, System.Drawing.FontStyle.Bold);
            this.btnReplay.Location = new System.Drawing.Point(150, 30);
            this.btnReplay.Name = "btnReplay";
            this.btnReplay.Size = new System.Drawing.Size(100, 35);
            this.btnReplay.TabIndex = 12;
            this.btnReplay.Text = "回放記錄";
            this.btnReplay.UseVisualStyleBackColor = true;
            this.btnReplay.Click += new System.EventHandler(this.btnReplay_Click);
            // 
            // btnClearRecord
            // 
            this.btnClearRecord.Enabled = false;
            this.btnClearRecord.Font = new System.Drawing.Font("Microsoft JhengHei", 10F);
            this.btnClearRecord.Location = new System.Drawing.Point(270, 30);
            this.btnClearRecord.Name = "btnClearRecord";
            this.btnClearRecord.Size = new System.Drawing.Size(100, 35);
            this.btnClearRecord.TabIndex = 13;
            this.btnClearRecord.Text = "清除記錄";
            this.btnClearRecord.UseVisualStyleBackColor = true;
            this.btnClearRecord.Click += new System.EventHandler(this.btnClearRecord_Click);
            // 
            // lstRecordedKeys
            // 
            this.lstRecordedKeys.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.lstRecordedKeys.FormattingEnabled = true;
            this.lstRecordedKeys.ItemHeight = 16;
            this.lstRecordedKeys.Location = new System.Drawing.Point(30, 100);
            this.lstRecordedKeys.Name = "lstRecordedKeys";
            this.lstRecordedKeys.Size = new System.Drawing.Size(390, 120);
            this.lstRecordedKeys.TabIndex = 14;
            // 
            // lblRecordedKeys
            // 
            this.lblRecordedKeys.AutoSize = true;
            this.lblRecordedKeys.Font = new System.Drawing.Font("Microsoft JhengHei", 10F);
            this.lblRecordedKeys.Location = new System.Drawing.Point(30, 80);
            this.lblRecordedKeys.Name = "lblRecordedKeys";
            this.lblRecordedKeys.Size = new System.Drawing.Size(93, 18);
            this.lblRecordedKeys.TabIndex = 15;
            this.lblRecordedKeys.Text = "記錄的按鍵:";
            // 
            // chkLoop
            // 
            this.chkLoop.AutoSize = true;
            this.chkLoop.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.chkLoop.Location = new System.Drawing.Point(380, 40);
            this.chkLoop.Name = "chkLoop";
            this.chkLoop.Size = new System.Drawing.Size(51, 20);
            this.chkLoop.TabIndex = 16;
            this.chkLoop.Text = "循環";
            this.chkLoop.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(680, 860);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnWorkflowMenu);
            this.Controls.Add(this.btnStart);
            this.Controls.Add(this.lblStartInfo);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.lblMode);
            this.Controls.Add(this.groupBoxHelp);
            this.Controls.Add(this.groupBoxTest);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "自動按鍵程式 - Auto Press v2.0";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBoxHelp.ResumeLayout(false);
            this.groupBoxTest.ResumeLayout(false);
            this.groupBoxTest.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

    private System.Windows.Forms.Button btnStart;
    private System.Windows.Forms.Label lblMode;
        private System.Windows.Forms.Label lblStartInfo;
    private System.Windows.Forms.Button btnWorkflowRun;
    private System.Windows.Forms.Button btnWorkflowRecord;
    private System.Windows.Forms.Button btnWorkflowMenu;
    // 已移除單鍵與間隔控制項
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox cmbApplications;
        private System.Windows.Forms.Label lblApplication;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnRecord;
        private System.Windows.Forms.Button btnReplay;
        private System.Windows.Forms.Button btnClearRecord;
        private System.Windows.Forms.ListBox lstRecordedKeys;
        private System.Windows.Forms.Label lblRecordedKeys;
    private System.Windows.Forms.CheckBox chkLoop;
    private System.Windows.Forms.GroupBox groupBoxHelp;
    private System.Windows.Forms.RichTextBox rtbHelp;
    private System.Windows.Forms.Button btnExport;
    private System.Windows.Forms.Button btnImport;
    private System.Windows.Forms.ComboBox cmbSpeed;
    private System.Windows.Forms.Label lblSpeed;
    private System.Windows.Forms.GroupBox groupBoxTest;
    private System.Windows.Forms.TextBox txtTest;
    private System.Windows.Forms.Button btnClearTest;
    private System.Windows.Forms.CheckBox chkTestMode;
    }
}

