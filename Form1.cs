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
using System.Diagnostics;

namespace AutoPressApp
{
    public partial class Form1 : Form
    {
        // Windows API for sending key presses
        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        // Windows API for window management
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        static extern bool IsWindowVisible(IntPtr hWnd);

        private const int KEYEVENTF_KEYUP = 0x0002;
        private const int SW_RESTORE = 9;
        
        private Timer autoTimer;
        private bool isRunning = false;
        private Keys selectedKey = Keys.F1; // Default key
        private IntPtr targetWindowHandle = IntPtr.Zero;
        private string targetWindowTitle = "";

        public Form1()
        {
            InitializeComponent();
            InitializeTimer();
            LoadRunningApplications();
        }

        private void InitializeTimer()
        {
            autoTimer = new Timer();
            autoTimer.Interval = 1000; // 1 second
            autoTimer.Tick += AutoTimer_Tick;
        }

        private void AutoTimer_Tick(object sender, EventArgs e)
        {
            if (targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle))
            {
                // Switch to target window and send key
                SetForegroundWindow(targetWindowHandle);
                ShowWindow(targetWindowHandle, SW_RESTORE);
                System.Threading.Thread.Sleep(50); // Small delay to ensure window is active
                
                SendKeyPress(selectedKey);
                UpdateStatus($"按鍵已發送至 [{targetWindowTitle}]: {selectedKey} - {DateTime.Now:HH:mm:ss}");
            }
            else
            {
                // Fallback to current window
                SendKeyPress(selectedKey);
                UpdateStatus($"按鍵已發送: {selectedKey} - {DateTime.Now:HH:mm:ss}");
            }
        }

        private void LoadRunningApplications()
        {
            cmbApplications.Items.Clear();
            cmbApplications.Items.Add("(當前焦點視窗)");
            
            Process[] processes = Process.GetProcesses();
            
            foreach (Process process in processes)
            {
                try
                {
                    if (process.MainWindowHandle != IntPtr.Zero && 
                        !string.IsNullOrEmpty(process.MainWindowTitle) &&
                        IsWindowVisible(process.MainWindowHandle))
                    {
                        string displayText = $"{process.ProcessName} - {process.MainWindowTitle}";
                        if (displayText.Length > 80)
                            displayText = displayText.Substring(0, 77) + "...";
                            
                        var item = new ApplicationItem
                        {
                            ProcessName = process.ProcessName,
                            WindowTitle = process.MainWindowTitle,
                            WindowHandle = process.MainWindowHandle,
                            DisplayText = displayText
                        };
                        
                        cmbApplications.Items.Add(item);
                    }
                }
                catch (Exception)
                {
                    // Skip processes that can't be accessed
                }
            }
            
            cmbApplications.SelectedIndex = 0; // Select "Current Focus Window" by default
        }

        public class ApplicationItem
        {
            public string ProcessName { get; set; }
            public string WindowTitle { get; set; }
            public IntPtr WindowHandle { get; set; }
            public string DisplayText { get; set; }
            
            public override string ToString()
            {
                return DisplayText;
            }
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

                // Get selected application
                if (cmbApplications.SelectedItem is ApplicationItem app)
                {
                    targetWindowHandle = app.WindowHandle;
                    targetWindowTitle = app.WindowTitle;
                }
                else
                {
                    targetWindowHandle = IntPtr.Zero;
                    targetWindowTitle = "當前焦點視窗";
                }

                autoTimer.Start();
                isRunning = true;
                btnStart.Text = "停止";
                UpdateStatus($"自動按鍵已開始 - 按鍵: {selectedKey}, 目標: {targetWindowTitle}");
            }
            else
            {
                autoTimer.Stop();
                isRunning = false;
                btnStart.Text = "開始";
                UpdateStatus("自動按鍵已停止");
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            LoadRunningApplications();
            UpdateStatus("應用程式清單已重新整理");
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
