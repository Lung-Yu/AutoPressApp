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
            
            // 檢查測試模式 checkbox 是否存在
            UpdateStatus($"[INIT] chkTestMode 控制項: {(chkTestMode != null ? "存在" : "不存在")}");
            if (chkTestMode != null)
            {
                UpdateStatus($"[INIT] chkTestMode 位置: ({chkTestMode.Left}, {chkTestMode.Top}), 大小: {chkTestMode.Width}x{chkTestMode.Height}");
                UpdateStatus($"[INIT] chkTestMode 可見: {chkTestMode.Visible}, 啟用: {chkTestMode.Enabled}");
            }
            
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
            UpdateStatus($"[TIMER] ReplayTimer_Tick 觸發, replayIndex={replayIndex}, recordedKeys.Count={recordedKeys.Count}");
            
            if (!isReplaying || replayIndex >= recordedKeys.Count)
            {
                // 結束回放
                UpdateStatus($"[TIMER] 回放結束條件: isReplaying={isReplaying}, replayIndex={replayIndex}");
                replayTimer.Stop();
                isReplaying = false;
                btnReplay.Text = "回放記錄";
                btnReplay.Enabled = true;
                UpdateStatus("回放完成");
                return;
            }
            
            var record = recordedKeys[replayIndex];
            bool testMode = chkTestMode != null && chkTestMode.Checked;
            
            UpdateStatus($"[TIMER] 處理第 {replayIndex + 1} 個按鍵: {record.Key}, 測試模式: {testMode}");
            
            if (testMode)
            {
                // 測試模式：寫入測試框
                if (txtTest != null)
                {
                    txtTest.AppendText(record.Key + " ");
                    UpdateStatus($"[TEST] 已寫入測試框: {record.Key}");
                }
                else
                {
                    UpdateStatus("[TEST] 錯誤: txtTest 控制項不存在");
                }
                UpdateStatus($"[測試模式] 已輸出: {record.Key} ({replayIndex + 1}/{recordedKeys.Count})");
            }
            else
            {
                // 正常模式：發送按鍵
                if (targetWindowHandle != IntPtr.Zero && IsWindow(targetWindowHandle))
                {
                    UpdateStatus($"[NORMAL] 設定目標視窗: {targetWindowHandle}");
                    SetForegroundWindow(targetWindowHandle);
                    ShowWindow(targetWindowHandle, SW_RESTORE);
                    System.Threading.Thread.Sleep(30);
                }
                else
                {
                    UpdateStatus("[NORMAL] 使用當前焦點視窗");
                }
                SendKeyPress(record.Key);
                UpdateStatus($"[正常模式] 已發送: {record.Key} ({replayIndex + 1}/{recordedKeys.Count})");
            }
            
            replayIndex++;
            
            // 設定下一個按鍵的間隔
            if (replayIndex < recordedKeys.Count)
            {
                int nextDelay = recordedKeys[replayIndex].DelayMs;
                int actualDelay = Math.Max(100, nextDelay);
                replayTimer.Interval = actualDelay;
                UpdateStatus($"[TIMER] 下一個間隔: {actualDelay}ms (原始: {nextDelay}ms)");
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
            UpdateStatus($"[LOG] SendKeyPress 開始發送: {key}");
            
            // 取得正確的虛擬鍵碼
            ushort virtualKey = GetVirtualKeyCode(key);
            
            // 使用 SendInput 發送
            var inputs = new List<INPUT>();
            inputs.Add(new INPUT
            {
                type = 1,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = virtualKey,
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
                        wVk = virtualKey,
                        wScan = 0,
                        dwFlags = KEYEVENTF_KEYUP,
                        time = 0,
                        dwExtraInfo = UIntPtr.Zero
                    }
                }
            });
            
            uint result = SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf<INPUT>());
            UpdateStatus($"[LOG] SendInput 結果: {result} (應該為 {inputs.Count}), VK={virtualKey:X2}");
            
            if (result == 0)
            {
                int error = Marshal.GetLastWin32Error();
                UpdateStatus($"[LOG] SendInput 失敗，錯誤代碼: {error}");
            }
        }

        private ushort GetVirtualKeyCode(Keys key)
        {
            // 處理特殊按鍵的對應關係
            switch (key)
            {
                case Keys.Space:
                    return 0x20; // VK_SPACE
                case Keys.Enter:
                    return 0x0D; // VK_RETURN
                case Keys.Escape:
                    return 0x1B; // VK_ESCAPE
                case Keys.Tab:
                    return 0x09; // VK_TAB
                case Keys.Back:
                    return 0x08; // VK_BACK
                case Keys.Delete:
                    return 0x2E; // VK_DELETE
                case Keys.Home:
                    return 0x24; // VK_HOME
                case Keys.End:
                    return 0x23; // VK_END
                case Keys.PageUp:
                    return 0x21; // VK_PRIOR
                case Keys.PageDown:
                    return 0x22; // VK_NEXT
                case Keys.Left:
                    return 0x25; // VK_LEFT
                case Keys.Up:
                    return 0x26; // VK_UP
                case Keys.Right:
                    return 0x27; // VK_RIGHT
                case Keys.Down:
                    return 0x28; // VK_DOWN
                case Keys.Insert:
                    return 0x2D; // VK_INSERT
                case Keys.F1:
                    return 0x70; // VK_F1
                case Keys.F2:
                    return 0x71; // VK_F2
                case Keys.F3:
                    return 0x72; // VK_F3
                case Keys.F4:
                    return 0x73; // VK_F4
                case Keys.F5:
                    return 0x74; // VK_F5
                case Keys.F6:
                    return 0x75; // VK_F6
                case Keys.F7:
                    return 0x76; // VK_F7
                case Keys.F8:
                    return 0x77; // VK_F8
                case Keys.F9:
                    return 0x78; // VK_F9
                case Keys.F10:
                    return 0x79; // VK_F10
                case Keys.F11:
                    return 0x7A; // VK_F11
                case Keys.F12:
                    return 0x7B; // VK_F12
                // 修飾鍵
                case Keys.LShiftKey:
                case Keys.RShiftKey:
                case Keys.ShiftKey:
                    return 0x10; // VK_SHIFT
                case Keys.LControlKey:
                case Keys.RControlKey:
                case Keys.ControlKey:
                    return 0x11; // VK_CONTROL
                case Keys.LMenu:
                case Keys.RMenu:
                case Keys.Menu:
                    return 0x12; // VK_MENU (Alt)
                case Keys.LWin:
                case Keys.RWin:
                    return 0x5B; // VK_LWIN
                // 數字鍵
                case Keys.D0:
                    return 0x30;
                case Keys.D1:
                    return 0x31;
                case Keys.D2:
                    return 0x32;
                case Keys.D3:
                    return 0x33;
                case Keys.D4:
                    return 0x34;
                case Keys.D5:
                    return 0x35;
                case Keys.D6:
                    return 0x36;
                case Keys.D7:
                    return 0x37;
                case Keys.D8:
                    return 0x38;
                case Keys.D9:
                    return 0x39;
                // 字母鍵 A-Z
                case Keys.A:
                    return 0x41;
                case Keys.B:
                    return 0x42;
                case Keys.C:
                    return 0x43;
                case Keys.D:
                    return 0x44;
                case Keys.E:
                    return 0x45;
                case Keys.F:
                    return 0x46;
                case Keys.G:
                    return 0x47;
                case Keys.H:
                    return 0x48;
                case Keys.I:
                    return 0x49;
                case Keys.J:
                    return 0x4A;
                case Keys.K:
                    return 0x4B;
                case Keys.L:
                    return 0x4C;
                case Keys.M:
                    return 0x4D;
                case Keys.N:
                    return 0x4E;
                case Keys.O:
                    return 0x4F;
                case Keys.P:
                    return 0x50;
                case Keys.Q:
                    return 0x51;
                case Keys.R:
                    return 0x52;
                case Keys.S:
                    return 0x53;
                case Keys.T:
                    return 0x54;
                case Keys.U:
                    return 0x55;
                case Keys.V:
                    return 0x56;
                case Keys.W:
                    return 0x57;
                case Keys.X:
                    return 0x58;
                case Keys.Y:
                    return 0x59;
                case Keys.Z:
                    return 0x5A;
                default:
                    // 對於其他按鍵，使用原本的轉換方式
                    return (ushort)key;
            }
        }

        private void UpdateStatus(string message)
        {
            if (lblStatus != null)
                lblStatus.Text = message;
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            UpdateStatus($"[LOG] btnStart_Click: isRunning={isRunning}");
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
            UpdateStatus("[LOG] BeginSequenceIfPossible 開始");
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
            UpdateStatus($"[LOG] 準備開始回放，測試模式: {testMode}，目標: {targetWindowTitle}");

            // 簡單的開始邏輯
            int firstDelay = recordedKeys[0].DelayMs <= 0 ? 100 : recordedKeys[0].DelayMs;
            replayTimer.Interval = firstDelay;
            replayTimer.Start();
            
            UpdateStatus($"[LOG] 定時器已啟動，間隔: {replayTimer.Interval}ms，啟用: {replayTimer.Enabled}");

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
            UpdateStatus("[BUTTON] btnReplay_Click 被點擊");
            
            // 確保不在記錄模式
            if (isRecording)
            {
                UpdateStatus("[BUTTON] 停止記錄模式");
                isRecording = false;
                btnRecord.Text = "開始記錄";
                btnRecord.BackColor = SystemColors.Control;
                btnRecord.ForeColor = SystemColors.ControlText;
            }
            
            if (recordedKeys.Count == 0)
            {
                UpdateStatus("[BUTTON] 沒有記錄的按鍵");
                MessageBox.Show("沒有記錄的按鍵可以回放！請先記錄一些按鍵。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            UpdateStatus($"[BUTTON] 準備回放 {recordedKeys.Count} 個按鍵");
            
            // 檢查測試模式狀態
            bool testMode = chkTestMode != null && chkTestMode.Checked;
            UpdateStatus($"[BUTTON] 測試模式檢查: chkTestMode={chkTestMode != null}, Checked={testMode}");
            if (chkTestMode != null)
            {
                UpdateStatus($"[BUTTON] chkTestMode 詳細: Visible={chkTestMode.Visible}, Enabled={chkTestMode.Enabled}, Text='{chkTestMode.Text}'");
            }
            
            // 獲取目標視窗
            if (cmbApplications.SelectedItem is ApplicationItem app)
            {
                targetWindowHandle = app.WindowHandle;
                targetWindowTitle = app.WindowTitle;
                UpdateStatus($"[BUTTON] 目標視窗: {targetWindowTitle} (Handle: {targetWindowHandle})");
            }
            else
            {
                targetWindowHandle = IntPtr.Zero;
                targetWindowTitle = "當前焦點視窗";
                UpdateStatus("[BUTTON] 使用當前焦點視窗");
            }
            
            // 開始簡單的單次回放
            replayIndex = 0;
            isReplaying = true;
            btnReplay.Text = "回放中...";
            btnReplay.Enabled = false;
            
            UpdateStatus($"[BUTTON] 設定回放狀態: replayIndex={replayIndex}, isReplaying={isReplaying}");
            
            // 使用簡單的定時器邏輯
            replayTimer.Interval = 100; // 100ms後開始第一個按鍵
            replayTimer.Start();
            
            UpdateStatus($"[BUTTON] 定時器已啟動: Interval={replayTimer.Interval}ms, Enabled={replayTimer.Enabled}");
            UpdateStatus($"開始回放，共 {recordedKeys.Count} 個按鍵");
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
