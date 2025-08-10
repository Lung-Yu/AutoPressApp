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
        // 使用 SendInput 取代過時的 keybd_event
        [StructLayout(LayoutKind.Sequential)]
        struct INPUT
        {
            public uint type; // 1 = Keyboard
            public InputUnion U;
        }

        [StructLayout(LayoutKind.Explicit)]
        struct InputUnion
        {
            [FieldOffset(0)] public KEYBDINPUT ki;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public UIntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

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
    private const int WM_KEYUP = 0x0101;
    private const int KEYEVENTF_KEYUP = 0x0002;
    private const int KEYEVENTF_SCANCODE = 0x0008;
        private const int SW_RESTORE = 9;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        
    private Timer replayTimer = new Timer();
    private bool isRunning = false;
        private bool isRecording = false;
        private bool isReplaying = false; // 用於判斷目前是否在回放
        private IntPtr targetWindowHandle = IntPtr.Zero;
        private string targetWindowTitle = "";
        
        // 按鍵記錄相關
        private List<KeyRecord> recordedKeys = new List<KeyRecord>();
        private DateTime lastKeyTime = DateTime.Now;
        private IntPtr keyboardHook = IntPtr.Zero;
        private LowLevelKeyboardProc hookProc;
    private int replayIndex = 0;
    // 速度倍率 (1.0 = 原始延遲) 可由 UI 調整（後續新增控制）
    private double delayMultiplier = 1.0;
    // 是否要循環回放（透過 UI CheckBox 控制）
    private bool LoopPlayback => chkLoop != null && chkLoop.Checked;

        public Form1()
        {
            InitializeComponent();
            InitializeTimers();
            LoadRunningApplications();
            
            // 初始化鍵盤鉤子
            hookProc = HookCallback;
            InstallGlobalHook();
            
            // 更新標籤顯示模式
            // 原 lblKey / 單鍵 UI 已移除

            // 使用說明
            if (rtbHelp != null)
            {
                rtbHelp.Text =
"【操作流程】\n" +
"1. 點『開始記錄』，輸入任意按鍵序列 (會記錄間隔)\n" +
"2. 再次點『開始記錄』= 停止記錄\n" +
"3. 勾選『循環』可無限回放\n" +
"4. 點『開始』= 回放記錄序列 (若無記錄會提示)\n" +
"5. 按 ESC 隨時緊急停止\n\n" +
"【熱鍵 / 控制】\n" +
"ESC : 停止所有執行 / 停止錄製\n" +
"開始記錄 : 進入錄製 / 停止錄製\n" +
"回放記錄 : 單次回放（若要無限請勾循環再用『開始』）\n" +
"清除記錄 : 清空已記錄序列\n\n" +
"【狀態說明】\n" +
"狀態列會顯示目前：錄製 / 回放第幾鍵 / 循環重新開始 等資訊\n" +
"目標視窗：可指定送出按鍵的應用程式 (未選擇則為目前焦點)";
            }

            // 加入快捷鍵提示 (新增的全域開始/停止)
            if (rtbHelp != null)
            {
                rtbHelp.AppendText("\nCtrl+Shift+S : 全域開始 (若有記錄)\nCtrl+Shift+X : 全域停止\n速度倍率: 可調整回放快慢 (預設1.0)\n");
            }
            
            // 註冊關閉事件
            this.FormClosing += Form1_FormClosing;
        }

        private void InitializeTimers()
        {
            replayTimer.Tick += ReplayTimer_Tick;
        }

        private void InstallGlobalHook()
        {
            if (keyboardHook == IntPtr.Zero)
            {
                using (Process curProcess = Process.GetCurrentProcess())
                using (ProcessModule curModule = curProcess.MainModule!)
                {
                    var moduleName = curModule.ModuleName ?? string.Empty;
                    keyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, hookProc, GetModuleHandle(moduleName), 0);
                }
            }
        }

        // 鍵盤鉤子回調函數
        private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN))
            {
                int vkCode = Marshal.ReadInt32(lParam);
                Keys key = (Keys)vkCode;
                // 判斷修飾鍵 (使用 GetAsyncKeyState)
                bool ctrl = IsKeyDown(Keys.LControlKey) || IsKeyDown(Keys.RControlKey) || IsKeyDown(Keys.ControlKey);
                bool shift = IsKeyDown(Keys.LShiftKey) || IsKeyDown(Keys.RShiftKey) || IsKeyDown(Keys.ShiftKey);

                if (ctrl && shift && key == Keys.S)
                {
                    BeginSequenceIfPossible();
                    return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
                }
                if (ctrl && shift && key == Keys.X)
                {
                    StopAll();
                    UpdateStatus("已透過全域停止熱鍵停止");
                    return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
                }

                if (key == Keys.Escape)
                {
                    // 停止回放/自動/記錄
                    StopAll();
                    UpdateStatus("已透過 ESC 停止所有執行");
                    return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
                }

                if (isRecording)
                {
                    // 記錄按鍵
                    DateTime now = DateTime.Now;
                    int delay = (int)(now - lastKeyTime).TotalMilliseconds;
                    if (recordedKeys.Count == 0) delay = 0;

                    var record = new KeyRecord
                    {
                        Key = key,
                        Timestamp = now,
                        DelayMs = delay
                    };

                    recordedKeys.Add(record);
                    lastKeyTime = now;
                    UpdateRecordedKeysList();
                    UpdateStatus($"記錄按鍵: {key} (已記錄 {recordedKeys.Count} 個按鍵)");
                }
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
            if (!isReplaying) { replayTimer.Stop(); return; }
            if (replayIndex >= recordedKeys.Count)
            {
                if (LoopPlayback && recordedKeys.Count > 0)
                {
                    replayIndex = 0; // 重頭
            int firstDelay = recordedKeys[0].DelayMs <= 0 ? 30 : recordedKeys[0].DelayMs;
            replayTimer.Interval = (int)Math.Max(1, firstDelay * delayMultiplier);
                    UpdateStatus($"循環回放重新開始 (共 {recordedKeys.Count} 按鍵)");
                    return; // 等下一輪 tick
                }
                // 單次結束
                replayTimer.Stop();
                isReplaying = false;
                btnReplay.Text = "回放記錄";
                btnReplay.Enabled = true;
                UpdateStatus("回放完成 (單次)");
                replayIndex = 0;
                return;
            }
            
            var record = recordedKeys[replayIndex];
            
            // 發送按鍵到目標視窗
            bool testMode = chkTestMode != null && chkTestMode.Checked;
            if (!testMode)
            {
                if (targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle))
                {
                    SetForegroundWindow(targetWindowHandle);
                    ShowWindow(targetWindowHandle, SW_RESTORE);
                    System.Threading.Thread.Sleep(30);
                }
                SendKeyPress(record.Key);
            }
            // 測試模式：僅寫入測試框，不實際送出
            if (testMode && txtTest != null)
            {
                txtTest.AppendText(record.Key + " ");
            }
            UpdateStatus($"回放按鍵 {replayIndex + 1}/{recordedKeys.Count}: {record.Key}");
            
            replayIndex++;
            
            // 設定下一個按鍵的延遲
            if (replayIndex < recordedKeys.Count)
            {
                int next = (int)(recordedKeys[replayIndex].DelayMs * delayMultiplier);
                replayTimer.Interval = Math.Max(1, next);
            }
        }

    // 已移除單鍵定時功能

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
            public string ProcessName { get; set; } = string.Empty;
            public string WindowTitle { get; set; } = string.Empty;
            public IntPtr WindowHandle { get; set; } = IntPtr.Zero;
            public string DisplayText { get; set; } = string.Empty;
            public override string ToString() => DisplayText;
        }

        private void SendKeyPress(Keys key)
        {
            // 使用 SendInput 發送
            var inputs = new List<INPUT>();
            inputs.Add(new INPUT
            {
                type = 1,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = (ushort)key,
                        wScan = 0,
                        dwFlags = 0,
                        time = 0,
                        dwExtraInfo = UIntPtr.Zero
                    }
                }
            });
            inputs.Add(new INPUT
            {
                type = 1,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = (ushort)key,
                        wScan = 0,
                        dwFlags = KEYEVENTF_KEYUP,
                        time = 0,
                        dwExtraInfo = UIntPtr.Zero
                    }
                }
            });
            SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf<INPUT>());
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
                BeginSequenceIfPossible();
            }
            else
            {
                StopAll();
                UpdateStatus("執行已停止");
            }
        }

        private void BeginSequenceIfPossible()
        {
            if (recordedKeys.Count == 0)
            {
                MessageBox.Show("尚未記錄任何按鍵，請先使用『開始記錄』功能。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

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

            replayIndex = 0;
            isReplaying = true;
            isRunning = true;
            if (btnStart != null) btnStart.Text = "停止";

            bool testMode = chkTestMode != null && chkTestMode.Checked;

            // 立即處理第一個按鍵（若其延遲為0），否則依延遲排程
            replayTimer.Stop();
            if (recordedKeys.Count > 0)
            {
                var first = recordedKeys[0];
                if (first.DelayMs == 0)
                {
                    if (testMode)
                    {
                        if (txtTest != null) txtTest.AppendText(first.Key + " ");
                    }
                    else
                    {
                        if (targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle))
                        {
                            SetForegroundWindow(targetWindowHandle);
                            ShowWindow(targetWindowHandle, SW_RESTORE);
                            System.Threading.Thread.Sleep(10);
                        }
                        SendKeyPress(first.Key);
                    }
                    replayIndex = 1; // 已處理第一個
                }
                int nextDelay;
                if (replayIndex < recordedKeys.Count)
                {
                    // 下一個按鍵的 delay 以其紀錄的 DelayMs (原指該鍵與前一鍵間隔)
                    nextDelay = recordedKeys[replayIndex].DelayMs;
                    if (replayIndex == 1 && first.DelayMs == 0)
                    {
                        // 已經立即送出第一鍵，此時 recordedKeys[1].DelayMs 仍是它與第一鍵的間隔，直接使用即可
                    }
                }
                else
                {
                    // 只有一個按鍵
                    nextDelay = 500; // fallback
                }
                replayTimer.Interval = Math.Max(1, (int)(nextDelay * delayMultiplier));
                replayTimer.Start();
            }
            UpdateStatus($"開始回放記錄序列{(testMode ? " (測試模式: 只顯示)" : "")}，共 {recordedKeys.Count} 個按鍵 → 目標: {(testMode ? "(不送出)" : targetWindowTitle)}");

            if (btnRecord != null) btnRecord.Enabled = false;
            if (btnReplay != null) btnReplay.Enabled = false;
            if (btnClearRecord != null) btnClearRecord.Enabled = false;
        }

        private void StopAll()
        {
            if (replayTimer.Enabled) replayTimer.Stop();
            isRunning = false;
            isReplaying = false;
            if (btnStart != null) btnStart.Text = "開始";
            if (btnReplay != null) btnReplay.Enabled = recordedKeys.Count > 0;
            if (btnClearRecord != null) btnClearRecord.Enabled = recordedKeys.Count > 0;
            if (btnRecord != null) btnRecord.Enabled = true;
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
                
                isRecording = true;
                btnRecord.Text = "停止記錄";
                btnRecord.BackColor = Color.Red;
                btnRecord.ForeColor = Color.White;
                UpdateStatus("開始記錄按鍵...請在任何地方按鍵");
            }
            else
            {
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
            
                // 單次回放使用與 BeginSequenceIfPossible 類似流程，但不改變 isRunning
                replayIndex = 0;
                isReplaying = true;
                btnReplay.Text = "回放中...";
                btnReplay.Enabled = false;
                bool testMode = chkTestMode != null && chkTestMode.Checked;
                replayTimer.Stop();
                if (recordedKeys[0].DelayMs == 0)
                {
                    // 立即發送第一鍵
                    if (testMode)
                    {
                        if (txtTest != null) txtTest.AppendText(recordedKeys[0].Key + " ");
                    }
                    else
                    {
                        if (targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle))
                        {
                            SetForegroundWindow(targetWindowHandle);
                            ShowWindow(targetWindowHandle, SW_RESTORE);
                            System.Threading.Thread.Sleep(10);
                        }
                        SendKeyPress(recordedKeys[0].Key);
                    }
                    replayIndex = 1;
                }
                int delay = 500;
                if (replayIndex < recordedKeys.Count)
                {
                    delay = recordedKeys[replayIndex].DelayMs;
                }
                replayTimer.Interval = Math.Max(1, (int)(delay * delayMultiplier));
                replayTimer.Start();
                var testSuffix = testMode ? " (測試模式: 只顯示)" : string.Empty;
                var targetLabel = testMode ? "(不送出)" : targetWindowTitle;
                UpdateStatus($"開始回放{testSuffix} 到 [{targetLabel}]，共 {recordedKeys.Count} 個按鍵");
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

    // 單鍵模式間隔事件已移除

        // 匯出記錄到 JSON
        private void ExportSequence(string filePath)
        {
            var simple = recordedKeys.Select(r => new { k = r.Key.ToString(), d = r.DelayMs }).ToList();
            var json = System.Text.Json.JsonSerializer.Serialize(simple, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
            System.IO.File.WriteAllText(filePath, json, Encoding.UTF8);
        }

        // 從 JSON 匯入記錄
        private void ImportSequence(string filePath)
        {
            if (!System.IO.File.Exists(filePath)) return;
            var json = System.IO.File.ReadAllText(filePath, Encoding.UTF8);
            var list = System.Text.Json.JsonSerializer.Deserialize<List<SequenceItem>>(json) ?? new();
            recordedKeys.Clear();
            int total = 0;
            foreach (var item in list)
            {
                if (Enum.TryParse<Keys>(item.k, out var key))
                {
                    recordedKeys.Add(new KeyRecord { Key = key, DelayMs = item.d, Timestamp = DateTime.Now.AddMilliseconds(total) });
                    total += item.d;
                }
            }
            UpdateRecordedKeysList();
            if (btnReplay != null) btnReplay.Enabled = recordedKeys.Count > 0;
            if (btnClearRecord != null) btnClearRecord.Enabled = recordedKeys.Count > 0;
            UpdateStatus($"已匯入 {recordedKeys.Count} 個按鍵");
        }

    private record SequenceItem(string k, int d);

    private void btnExport_Click(object sender, EventArgs e)
    {
        if (recordedKeys.Count == 0)
        {
            MessageBox.Show("目前沒有記錄可匯出", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        using var sfd = new SaveFileDialog { Filter = "JSON (*.json)|*.json", FileName = "sequence.json" };
        if (sfd.ShowDialog() == DialogResult.OK)
        {
            try
            {
                ExportSequence(sfd.FileName);
                UpdateStatus("匯出完成: " + sfd.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("匯出失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private void btnImport_Click(object sender, EventArgs e)
    {
        using var ofd = new OpenFileDialog { Filter = "JSON (*.json)|*.json" };
        if (ofd.ShowDialog() == DialogResult.OK)
        {
            try
            {
                ImportSequence(ofd.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("匯入失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private void cmbSpeed_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (sender is ComboBox cb && cb.SelectedItem is string txt)
        {
            if (txt.EndsWith("x") && double.TryParse(txt.TrimEnd('x'), out var v))
            {
                delayMultiplier = 1.0 / v; // 播放速度倍率 => 延遲縮放
                UpdateStatus($"速度倍率設定: {txt}");
            }
        }
    }

    private void btnClearTest_Click(object sender, EventArgs e)
    {
        if (txtTest != null) txtTest.Clear();
    }

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);
    private static bool IsKeyDown(Keys key) => (GetAsyncKeyState((int)key) & 0x8000) != 0;
    }
}
