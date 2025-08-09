using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace AutoPressApp
{
    public partial class Form1 : Form
    {
        // Windows API for sending key presses
        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        private const int KEYEVENTF_KEYUP = 0x0002;
        private Timer autoTimer;
        private bool isRunning = false;
        private Keys selectedKey = Keys.F1; // Default key

        public Form1()
        {
            InitializeComponent();
            InitializeTimer();
        }

        private void InitializeTimer()
        {
            autoTimer = new Timer();
            autoTimer.Interval = 1000; // 1 second
            autoTimer.Tick += AutoTimer_Tick;
        }

        private void AutoTimer_Tick(object sender, EventArgs e)
        {
            SendKeyPress(selectedKey);
            UpdateStatus($"按鍵已發送: {selectedKey} - {DateTime.Now:HH:mm:ss}");
        }

        private void SendKeyPress(Keys key)
        {
            byte vkCode = (byte)key;
            keybd_event(vkCode, 0, 0, UIntPtr.Zero);           // Key down
            keybd_event(vkCode, 0, KEYEVENTF_KEYUP, UIntPtr.Zero); // Key up
        }

        private void UpdateStatus(string message)
        {
            if (lblStatus != null)
                lblStatus.Text = message;
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (!isRunning)
            {
                // Get selected key from combo box
                if (cmbKeys.SelectedItem != null)
                {
                    Enum.TryParse(cmbKeys.SelectedItem.ToString(), out selectedKey);
                }

                autoTimer.Start();
                isRunning = true;
                btnStart.Text = "停止";
                UpdateStatus($"自動按鍵已開始 - 按鍵: {selectedKey}");
            }
            else
            {
                autoTimer.Stop();
                isRunning = false;
                btnStart.Text = "開始";
                UpdateStatus("自動按鍵已停止");
            }
        }

        private void numInterval_ValueChanged(object sender, EventArgs e)
        {
            if (autoTimer != null)
            {
                autoTimer.Interval = (int)(numInterval.Value * 1000);
                UpdateStatus($"間隔已設定為 {numInterval.Value} 秒");
            }
        }
    }
}
