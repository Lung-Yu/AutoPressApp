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
            this.cmbKeys = new System.Windows.Forms.ComboBox();
            this.lblKey = new System.Windows.Forms.Label();
            this.lblInterval = new System.Windows.Forms.Label();
            this.numInterval = new System.Windows.Forms.NumericUpDown();
            this.lblStatus = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cmbApplications = new System.Windows.Forms.ComboBox();
            this.lblApplication = new System.Windows.Forms.Label();
            this.btnRefresh = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.numInterval)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnStart
            // 
            this.btnStart.Font = new System.Drawing.Font("Microsoft JhengHei", 12F, System.Drawing.FontStyle.Bold);
            this.btnStart.Location = new System.Drawing.Point(30, 280);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(100, 40);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "開始";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // cmbKeys
            // 
            this.cmbKeys.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbKeys.Font = new System.Drawing.Font("Microsoft JhengHei", 10F);
            this.cmbKeys.FormattingEnabled = true;
            this.cmbKeys.Items.AddRange(new object[] {
            "F1",
            "F2",
            "F3",
            "F4",
            "F5",
            "Space",
            "Enter",
            "A",
            "B",
            "C",
            "D",
            "E"});
            this.cmbKeys.Location = new System.Drawing.Point(100, 30);
            this.cmbKeys.Name = "cmbKeys";
            this.cmbKeys.Size = new System.Drawing.Size(121, 25);
            this.cmbKeys.TabIndex = 1;
            this.cmbKeys.SelectedIndex = 0;
            // 
            // lblKey
            // 
            this.lblKey.AutoSize = true;
            this.lblKey.Font = new System.Drawing.Font("Microsoft JhengHei", 10F);
            this.lblKey.Location = new System.Drawing.Point(30, 33);
            this.lblKey.Name = "lblKey";
            this.lblKey.Size = new System.Drawing.Size(65, 18);
            this.lblKey.TabIndex = 2;
            this.lblKey.Text = "選擇按鍵:";
            // 
            // lblInterval
            // 
            this.lblInterval.AutoSize = true;
            this.lblInterval.Font = new System.Drawing.Font("Microsoft JhengHei", 10F);
            this.lblInterval.Location = new System.Drawing.Point(30, 73);
            this.lblInterval.Name = "lblInterval";
            this.lblInterval.Size = new System.Drawing.Size(65, 18);
            this.lblInterval.TabIndex = 3;
            this.lblInterval.Text = "間隔(秒):";
            // 
            // numInterval
            // 
            this.numInterval.DecimalPlaces = 1;
            this.numInterval.Font = new System.Drawing.Font("Microsoft JhengHei", 10F);
            this.numInterval.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.numInterval.Location = new System.Drawing.Point(100, 70);
            this.numInterval.Maximum = new decimal(new int[] {
            60,
            0,
            0,
            0});
            this.numInterval.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.numInterval.Name = "numInterval";
            this.numInterval.Size = new System.Drawing.Size(120, 25);
            this.numInterval.TabIndex = 4;
            this.numInterval.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numInterval.ValueChanged += new System.EventHandler(this.numInterval_ValueChanged);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Font = new System.Drawing.Font("Microsoft JhengHei", 9F);
            this.lblStatus.ForeColor = System.Drawing.Color.Blue;
            this.lblStatus.Location = new System.Drawing.Point(30, 340);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(67, 16);
            this.lblStatus.TabIndex = 5;
            this.lblStatus.Text = "狀態: 待命中";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnRefresh);
            this.groupBox1.Controls.Add(this.lblApplication);
            this.groupBox1.Controls.Add(this.cmbApplications);
            this.groupBox1.Controls.Add(this.lblKey);
            this.groupBox1.Controls.Add(this.cmbKeys);
            this.groupBox1.Controls.Add(this.lblInterval);
            this.groupBox1.Controls.Add(this.numInterval);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft JhengHei", 10F);
            this.groupBox1.Location = new System.Drawing.Point(30, 30);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(450, 200);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "設定";
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
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(520, 380);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnStart);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "自動按鍵程式 - Auto Press";
            ((System.ComponentModel.ISupportInitialize)(this.numInterval)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnStart;
        private System.Windows.Forms.ComboBox cmbKeys;
        private System.Windows.Forms.Label lblKey;
        private System.Windows.Forms.Label lblInterval;
        private System.Windows.Forms.NumericUpDown numInterval;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox cmbApplications;
        private System.Windows.Forms.Label lblApplication;
        private System.Windows.Forms.Button btnRefresh;
    }
}

