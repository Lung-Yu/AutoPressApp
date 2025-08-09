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
    // 按鍵記錄結構
    public class KeyRecord
    {
        public Keys Key { get; set; }
        public DateTime Timestamp { get; set; }
        public int DelayMs { get; set; }
        
        public override string ToString()
        {
            return $"{Key} (延遲: {DelayMs}ms)";
        }
    }

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

        // 低階鍵盤鉤子用於記錄按鍵
        [DllImport("user32.dll")]
        static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        static extern IntPtr GetModuleHandle(string lpModuleName);

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int KEYEVENTF_KEYUP = 0x0002;
        private const int SW_RESTORE = 9;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        
        private Timer autoTimer;
        private Timer replayTimer;
        private bool isRunning = false;
        private bool isRecording = false;
        private bool isReplaying = false;
        private Keys selectedKey = Keys.F1;
        private IntPtr targetWindowHandle = IntPtr.Zero;
        private string targetWindowTitle = "";
        
        // 按鍵記錄相關
        private List<KeyRecord> recordedKeys = new List<KeyRecord>();
        private DateTime lastKeyTime = DateTime.Now;
        private IntPtr keyboardHook = IntPtr.Zero;
        private LowLevelKeyboardProc hookProc;
        private int replayIndex = 0;

        public Form1()
        {
            InitializeComponent();
            InitializeTimers();
            LoadRunningApplications();
            
            // 初始化鍵盤鉤子
            hookProc = HookCallback;
            
            // 註冊關閉事件
            this.FormClosing += Form1_FormClosing;
        }

        private void InitializeTimers()
        {
            // 原有的自動按鍵定時器
            autoTimer = new Timer();
            autoTimer.Interval = 1000; // 1 second
            autoTimer.Tick += AutoTimer_Tick;
            
            // 新增的回放定時器
            replayTimer = new Timer();
            replayTimer.Tick += ReplayTimer_Tick;
        }

        // 鍵盤鉤子回調函數
        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN && isRecording)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                Keys key = (Keys)vkCode;
                
                // 記錄按鍵
                DateTime now = DateTime.Now;
                int delay = (int)(now - lastKeyTime).TotalMilliseconds;
                
                if (recordedKeys.Count == 0)
                    delay = 0; // 第一個按鍵沒有延遲
                
                var record = new KeyRecord
                {
                    Key = key,
                    Timestamp = now,
                    DelayMs = delay
                };
                
                recordedKeys.Add(record);
                lastKeyTime = now;
                
                // 更新顯示
                UpdateRecordedKeysList();
                UpdateStatus($"記錄按鍵: {key} (已記錄 {recordedKeys.Count} 個按鍵)");
            }
            
            return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
        }

        private void UpdateRecordedKeysList()
        {
            if (lstRecordedKeys.InvokeRequired)
            {
                lstRecordedKeys.Invoke(new Action(UpdateRecordedKeysList));
                return;
            }
            
            lstRecordedKeys.Items.Clear();
            for (int i = 0; i < recordedKeys.Count; i++)
            {
                lstRecordedKeys.Items.Add($"{i + 1}. {recordedKeys[i]}");
            }
        }

        private void ReplayTimer_Tick(object sender, EventArgs e)
        {
            if (replayIndex >= recordedKeys.Count)
            {
                // 回放完成
                replayTimer.Stop();
                isReplaying = false;
                btnReplay.Text = "回放記錄";
                btnReplay.Enabled = true;
                UpdateStatus("回放完成");
                replayIndex = 0;
                return;
            }
            
            var record = recordedKeys[replayIndex];
            
            // 發送按鍵到目標視窗
            if (targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle))
            {
                SetForegroundWindow(targetWindowHandle);
                ShowWindow(targetWindowHandle, SW_RESTORE);
                System.Threading.Thread.Sleep(50);
            }
            
            SendKeyPress(record.Key);
            UpdateStatus($"回放按鍵 {replayIndex + 1}/{recordedKeys.Count}: {record.Key}");
            
            replayIndex++;
            
            // 設定下一個按鍵的延遲
            if (replayIndex < recordedKeys.Count)
            {
                replayTimer.Interval = recordedKeys[replayIndex].DelayMs;
            }
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

        private void btnRecord_Click(object sender, EventArgs e)
        {
            if (!isRecording)
            {
                // 開始記錄
                recordedKeys.Clear();
                UpdateRecordedKeysList();
                lastKeyTime = DateTime.Now;
                
                // 安裝鍵盤鉤子
                using (Process curProcess = Process.GetCurrentProcess())
                using (ProcessModule curModule = curProcess.MainModule)
                {
                    keyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, hookProc,
                        GetModuleHandle(curModule.ModuleName), 0);
                }
                
                isRecording = true;
                btnRecord.Text = "停止記錄";
                btnRecord.BackColor = Color.Red;
                btnRecord.ForeColor = Color.White;
                UpdateStatus("開始記錄按鍵...請在任何地方按鍵");
            }
            else
            {
                // 停止記錄
                if (keyboardHook != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(keyboardHook);
                    keyboardHook = IntPtr.Zero;
                }
                
                isRecording = false;
                btnRecord.Text = "開始記錄";
                btnRecord.BackColor = SystemColors.Control;
                btnRecord.ForeColor = SystemColors.ControlText;
                UpdateStatus($"記錄完成，共記錄 {recordedKeys.Count} 個按鍵");
                
                btnReplay.Enabled = recordedKeys.Count > 0;
                btnClearRecord.Enabled = recordedKeys.Count > 0;
            }
        }

        private void btnReplay_Click(object sender, EventArgs e)
        {
            if (recordedKeys.Count == 0)
            {
                MessageBox.Show("沒有記錄的按鍵可以回放！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            // 獲取目標視窗
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
            
            // 開始回放
            replayIndex = 0;
            isReplaying = true;
            btnReplay.Text = "回放中...";
            btnReplay.Enabled = false;
            
            // 設定第一個按鍵的延遲（通常為0）
            replayTimer.Interval = recordedKeys[0].DelayMs > 0 ? recordedKeys[0].DelayMs : 100;
            replayTimer.Start();
            
            UpdateStatus($"開始回放到 [{targetWindowTitle}]，共 {recordedKeys.Count} 個按鍵");
        }

        private void btnClearRecord_Click(object sender, EventArgs e)
        {
            recordedKeys.Clear();
            UpdateRecordedKeysList();
            btnReplay.Enabled = false;
            btnClearRecord.Enabled = false;
            UpdateStatus("記錄已清除");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (keyboardHook != IntPtr.Zero)
            {
                UnhookWindowsHookEx(keyboardHook);
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
