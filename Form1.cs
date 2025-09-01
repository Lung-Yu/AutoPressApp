using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using AutoPressApp.Services;
using AutoPressApp.Models;

namespace AutoPressApp
{
    // (已移除未使用的 KeyRecord 結構)

    public partial class Form1 : Form
    {
        private void btnDeleteStep_Click(object sender, EventArgs e)
        {
            if (lstRecordedKeys == null || liveWorkflowSteps == null) return;
            var indices = lstRecordedKeys.SelectedIndices;
            if (indices.Count == 0) return;
            // 先將 index 由大到小排序，避免移除時 index 錯亂
            var toRemove = new List<int>();
            foreach (int idx in indices) toRemove.Add(idx);
            toRemove.Sort((a, b) => b.CompareTo(a));
            foreach (int idx in toRemove)
            {
                if (idx >= 0 && idx < liveWorkflowSteps.Count)
                {
                    liveWorkflowSteps.RemoveAt(idx);
                    lstRecordedKeys.Items.RemoveAt(idx);
                }
            }
            // 更新按鈕狀態
            btnReplay.Enabled = liveWorkflowSteps.Count > 0;
            btnClearRecord.Enabled = liveWorkflowSteps.Count > 0;
            UpdateStatus($"已刪除 {toRemove.Count} 個步驟");
        }
    // ---- Win32 Interop ----
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

        // 添加 PostMessage API 作為備用方案
        [DllImport("user32.dll")]
        static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        // 添加 SendMessage API 用於直接字符發送
        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        // 添加 keybd_event API 作為備用方案
        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        // 添加更多API用於診斷
        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("kernel32.dll")]
        static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;
    private const int WM_KEYUP = 0x0101;
    private const int KEYEVENTF_KEYUP = 0x0002;
        private const int SW_RESTORE = 9;

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        
    private bool isRunning = false;
    private IntPtr targetWindowHandle = IntPtr.Zero;
    // Legacy 相關欄位已移除
    private IntPtr keyboardHook = IntPtr.Zero;
    private LowLevelKeyboardProc hookProc;
    private double delayMultiplier = 1.0;
    // 防止錄製 finalize 重入 (0 = 未完成, 1 = 已完成)
    private int recordingFinalizeState = 0;

        // 工作流程相關
        private CancellationTokenSource? workflowCts;
        private RecorderService? recorder;
    private List<AutoPressApp.Steps.Step>? liveWorkflowSteps;

    // 已移除 workflowMenu / btnWorkflowMenu
        // 模式狀態 (統一顯示 Recording / Playing)
    private enum RunMode { Idle, RecordingWorkflow, PlayingWorkflow }
        private RunMode currentMode = RunMode.Idle;
        private void SetMode(RunMode mode)
        {
            currentMode = mode;
            try
            {
                if (InvokeRequired)
                {
                    Invoke(new Action<RunMode>(SetMode), mode);
                    return;
                }
                string txt = mode switch
                {
                    RunMode.Idle => "狀態: 空閒",
                    RunMode.RecordingWorkflow => "狀態: 錄製流程中 (ESC 停止)",
                    RunMode.PlayingWorkflow => "狀態: 回放流程中 (ESC 停止)",
                    _ => "狀態: ?"
                };
                if (lblMode != null) lblMode.Text = txt;
                Color back; Color fore;
                switch (mode)
                {
            case RunMode.RecordingWorkflow:
                        back = Color.Red; fore = Color.White; break;
            case RunMode.PlayingWorkflow:
                        back = Color.ForestGreen; fore = Color.White; break;
                    default:
                        back = SystemColors.ControlLight; fore = Color.Black; break;
                }
                if (lblMode != null)
                {
                    lblMode.BackColor = back;
                    lblMode.ForeColor = fore;
                }
                if (btnStart != null)
                {
            if (mode == RunMode.PlayingWorkflow)
                    {
                        btnStart.Text = "停止";
                        btnStart.Enabled = true;
                        btnStart.BackColor = back;
                        btnStart.ForeColor = fore;
                    }
                    else if (mode == RunMode.Idle)
                    {
                        btnStart.Text = "開始";
                        btnStart.Enabled = true;
                        btnStart.BackColor = SystemColors.Control;
                        btnStart.ForeColor = SystemColors.ControlText;
                    }
                    else
                    {
                        btnStart.Text = "開始";
                        btnStart.Enabled = false; // 錄製期間鎖定開始
                        btnStart.BackColor = SystemColors.Control;
                        btnStart.ForeColor = SystemColors.ControlText;
                    }
                }
                if (lblRecordedKeys != null)
                {
                    if (mode == RunMode.RecordingWorkflow || mode == RunMode.PlayingWorkflow || (liveWorkflowSteps != null && liveWorkflowSteps.Count > 0))
                        lblRecordedKeys.Text = "錄製的流程步驟:";
                    else
                        lblRecordedKeys.Text = "記錄的按鍵:";
                }
            }
            catch { }
        }

        public Form1()
        {
            InitializeComponent();
            LoadRunningApplications();
            LoadSavedWorkflowList();
            
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
"【快捷鍵總覽】\n" +
"Ctrl+Shift+R : 錄製流程 (滑鼠+視窗+組合鍵)\n" +
"Ctrl+Shift+W : 選擇並執行 Workflow JSON\n" +
"Ctrl+Shift+S : 開始播放 (優先流程步驟, 無則舊按鍵)\n" +
"Ctrl+Shift+X / ESC : 停止所有執行\n" +
"\n【開始按鈕行為】\n" +
"有錄製的流程步驟 -> 直接執行該流程\n" +
"沒有流程步驟但有舊按鍵序列 -> 執行舊按鍵回放\n" +
"都沒有 -> 提示需要先錄製/記錄\n" +
"\n【流程操作】\n" +
"1. 『流程功能』或 Ctrl+Shift+R 開始錄製\n" +
"2. 再次 Ctrl+Shift+R 停止 -> 可選預覽與儲存\n" +
"3. 『流程功能』> 執行 或 Ctrl+Shift+W 執行 JSON\n" +
"4. 『重新預覽最近錄製』快速重放當前步驟集合\n" +
"\n【錄製內容】FocusWindow / Click / KeyCombo / Delay 會即時列出供檢閱。Delay=事件間隔。\n" +
"\n【提示】如組合鍵未正常回放可嘗試系統管理員執行。";
            }

            // 加入快捷鍵提示 (新增的全域開始/停止)
            if (rtbHelp != null)
            {
                rtbHelp.AppendText("\nCtrl+Shift+S : 全域開始 (優先流程步驟)\nCtrl+Shift+X / ESC : 全域停止\nCtrl+Shift+W : 選擇並執行 Workflow JSON\nCtrl+Shift+R : 開始/停止 錄製工作流程\n速度倍率: 可調整回放快慢 (預設1.0)\n提示: 錄製中『開始』與 Ctrl+Shift+S 會被鎖定，需先停止錄製。\n");
            }
            
            // 註冊關閉事件
            this.FormClosing += Form1_FormClosing;
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
                    // 全域開始：優先執行已錄製流程步驟，其次舊按鍵序列
                    StartPreferred();
                    return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
                }
                if (ctrl && shift && key == Keys.W)
                {
                    // Load and run a workflow JSON
                    RunWorkflowViaPicker();
                    return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
                }
                if (ctrl && shift && key == Keys.R)
                {
                    // 播放期間忽略錄製熱鍵，避免中途插入新錄製
                    if (currentMode == RunMode.PlayingWorkflow)
                    {
                        UpdateStatus("[Hotkey] 播放中忽略錄製熱鍵");
                    }
                    else
                    {
                        ToggleRecordWorkflow();
                    }
                    return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
                }
                if (ctrl && shift && key == Keys.P)
                {
                    if (recorder == null && liveWorkflowSteps != null && liveWorkflowSteps.Count > 0)
                        StartPreferred();
                    else
                        UpdateStatus("[Workflow] 尚無流程步驟可播放，請先 Ctrl+Shift+R 錄製");
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
                    if (recorder != null && currentMode == RunMode.RecordingWorkflow)
                    {
                        // 交由 RecorderService 內部 ESC 偵測觸發 StopToWorkflow -> OnStopped -> Finalize
                        UpdateStatus("[Recorder] ESC 停止錄製中...");
                    }
                    else
                    {
                        StopAll();
                        UpdateStatus("已透過 ESC 停止所有執行");
                    }
                    return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
                }

            }

            return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
        }

        // Legacy UpdateRecordedKeysList method removed

    // ...existing code...

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
                        
                        // 為記事本添加特別的日誌
                        if (process.ProcessName.ToLower().Contains("notepad"))
                        {
                            UpdateStatus($"[LOAD] 發現記事本: {process.ProcessName}, Handle: {process.MainWindowHandle}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Skip processes that can't be accessed
                    UpdateStatus($"[LOAD] 跳過程序 {process.ProcessName}: {ex.Message}");
                }
            }
            
            cmbApplications.SelectedIndex = 0; // Select "Current Focus Window" by default
            UpdateStatus($"[LOAD] 載入了 {cmbApplications.Items.Count - 1} 個可用視窗");
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
            
            // 如果有指定目標視窗，先將其設為焦點
            if (targetWindowHandle != IntPtr.Zero)
            {
                UpdateStatus($"[LOG] 檢查目標視窗: {targetWindowHandle}");
                bool isValidWindow = IsWindow(targetWindowHandle);
                bool isVisible = IsWindowVisible(targetWindowHandle);
                UpdateStatus($"[LOG] 視窗狀態: Valid={isValidWindow}, Visible={isVisible}");
                
                if (isValidWindow && isVisible)
                {
                    // 設置為前景視窗
                    bool setForeground = SetForegroundWindow(targetWindowHandle);
                    UpdateStatus($"[LOG] SetForegroundWindow 結果: {setForeground}");
                    
                    // 確保視窗可見
                    ShowWindow(targetWindowHandle, SW_RESTORE);
                    
                    // 等待視窗切換完成
                    System.Threading.Thread.Sleep(200);
                    
                    // 確認前景視窗
                    IntPtr foregroundWindow = GetForegroundWindow();
                    bool isForeground = foregroundWindow == targetWindowHandle;
                    UpdateStatus($"[LOG] 前景視窗檢查: Expected={targetWindowHandle}, Actual={foregroundWindow}, Match={isForeground}");
                }
                else
                {
                    UpdateStatus("[LOG] 警告: 目標視窗無效，使用當前焦點視窗");
                    targetWindowHandle = IntPtr.Zero;
                }
            }
            else
            {
                UpdateStatus("[LOG] 沒有指定目標視窗，使用當前焦點視窗");
            }
            
            // 發送按鍵到系統
            SendInputToSystem(key);
        }
        
        private void SendInputToSystem(Keys key)
        {
            // 取得正確的虛擬鍵碼
            ushort virtualKey = GetVirtualKeyCode(key);
            UpdateStatus($"[LOG] 虛擬鍵碼轉換: {key} -> 0x{virtualKey:X2}");
            
            // 使用 SendInput 發送到系統
            var inputs = new List<INPUT>();
            
            // 按下按鍵
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
            
            // 釋放按鍵
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
            
            UpdateStatus($"[LOG] 準備發送 {inputs.Count} 個全域輸入事件");
            
            uint result = SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf<INPUT>());
            UpdateStatus($"[LOG] 全域 SendInput 結果: {result}/{inputs.Count}, VK=0x{virtualKey:X2}");
            
            if (result == 0)
            {
                int error = Marshal.GetLastWin32Error();
                UpdateStatus($"[LOG] 全域 SendInput 失敗，錯誤代碼: {error}");
                
                // 嘗試使用 keybd_event 作為備用
                UpdateStatus("[LOG] 嘗試使用 keybd_event 備用方案");
                keybd_event((byte)virtualKey, 0, 0, UIntPtr.Zero); // 按下
                System.Threading.Thread.Sleep(10);
                keybd_event((byte)virtualKey, 0, KEYEVENTF_KEYUP, UIntPtr.Zero); // 釋放
                UpdateStatus($"[LOG] keybd_event 完成: {key}");
            }
            else if (result != inputs.Count)
            {
                UpdateStatus($"[LOG] 警告: SendInput 只處理了 {result}/{inputs.Count} 個事件");
            }
            else
            {
                UpdateStatus($"[LOG] 全域 SendInput 成功發送按鍵: {key}");
            }
        }

        private char GetCharFromKey(Keys key)
        {
            // 將常見的按鍵轉換為對應的字符
            if (key >= Keys.A && key <= Keys.Z)
            {
                return (char)('A' + (key - Keys.A));
            }
            
            if (key >= Keys.D0 && key <= Keys.D9)
            {
                return (char)('0' + (key - Keys.D0));
            }
            
            switch (key)
            {
                case Keys.Space: return ' ';
                case Keys.Enter: return '\r';
                case Keys.Tab: return '\t';
                default: return '\0'; // 無法轉換
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

    // (已移除 TestSendInput / 直接測試用程式碼)

        private void UpdateStatus(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(UpdateStatus), message);
                return;
            }
            
            // 更新狀態標籤
            if (lblStatus != null)
                lblStatus.Text = message;
                
            // 同時將訊息添加到底部的日誌文字框
            if (txtTest != null)
            {
                txtTest.AppendText($"{DateTime.Now:HH:mm:ss} - {message}{Environment.NewLine}");
                txtTest.SelectionStart = txtTest.Text.Length;
                txtTest.ScrollToCaret();
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            // 若目前在任一播放模式，轉成停止
            if (currentMode == RunMode.PlayingWorkflow || isRunning)
            {
                UpdateStatus("[UI] 停止按鈕觸發");
                StopAll();
                return;
            }
            StartPreferred();
        }

        private void StartPreferred()
        {
            UpdateStatus($"[LOG] StartPreferred: isRunning={isRunning}");
            if (isRunning)
            {
                StopAll();
                UpdateStatus("執行已停止");
                return;
            }
            // 正在錄製時不可啟動
            if (recorder != null)
            {
                UpdateStatus("[Start] 正在錄製中，無法啟動播放");
                return;
            }
            if (liveWorkflowSteps != null && liveWorkflowSteps.Count > 0)
            {
                UpdateStatus("[Start] 執行現有流程步驟");
                isRunning = true;
                _ = RunWorkflowPreviewAsync(new Workflow { Name = "Live Workflow", Steps = new List<AutoPressApp.Steps.Step>(liveWorkflowSteps) });
                return;
            }
            MessageBox.Show("尚未錄製任何流程步驟。請先錄製流程。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            UpdateStatus("[Start] 尚未錄製任何流程步驟，請先錄製流程。");
        }

        private void StopAll()
        {
            UpdateStatus("[LOG] StopAll 被呼叫");
            isRunning = false;
            workflowCts?.Cancel();
            if (recorder != null)
            {
                if (currentMode == RunMode.RecordingWorkflow)
                {
                    FinalizeWorkflowRecording(viaEsc: false, interactive: false); // 保留步驟
                }
                else
                {
                    try
                    {
                        recorder.OnLog -= UpdateStatus;
                        recorder.StepCaptured -= Recorder_StepCaptured;
                        recorder.Dispose();
                    }
                    catch { }
                    recorder = null;
                    UpdateStatus("[Recorder] 錄製已中止");
                }
            }
            // (已移除舊 isRecording 流程)
            if (btnReplay != null)
            {
                btnReplay.Text = "回放記錄";
                btnReplay.Enabled = liveWorkflowSteps != null && liveWorkflowSteps.Count > 0;
            }
            if (btnClearRecord != null) btnClearRecord.Enabled = liveWorkflowSteps != null && liveWorkflowSteps.Count > 0;
            if (btnRecord != null) btnRecord.Enabled = true;
            SetMode(RunMode.Idle);
            UpdateStatus("[LOG] 所有狀態已重置");
        }

        // ===== Saved Workflow Management =====
        private void LoadSavedWorkflowList()
        {
            try
            {
                if (cmbSavedWorkflows == null) return;
                var items = WorkflowStorage.List();
                cmbSavedWorkflows.Items.Clear();
                foreach (var info in items)
                {
                    cmbSavedWorkflows.Items.Add(info.Name);
                }
                if (cmbSavedWorkflows.Items.Count > 0) cmbSavedWorkflows.SelectedIndex = 0;
                UpdateStatus($"[SavedWF] 已載入 {cmbSavedWorkflows.Items.Count} 個已儲存流程");
            }
            catch (Exception ex)
            {
                UpdateStatus("[SavedWF] 載入清單失敗: " + ex.Message);
            }
        }

        private void cmbSavedWorkflows_SelectedIndexChanged(object? sender, EventArgs e)
        {
            // 僅選擇，不自動載入執行
        }

        private void btnRefreshSaved_Click(object? sender, EventArgs e)
        {
            LoadSavedWorkflowList();
        }

        private void btnLoadSaved_Click(object? sender, EventArgs e)
        {
            if (cmbSavedWorkflows == null || cmbSavedWorkflows.SelectedItem == null) return;
            string name = cmbSavedWorkflows.SelectedItem.ToString()!;
            var wf = WorkflowStorage.Load(name);
            if (wf == null)
            {
                MessageBox.Show("載入失敗或檔案不存在", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            liveWorkflowSteps = new List<AutoPressApp.Steps.Step>(wf.Steps);
            RebuildTreeView();
            // 同步 Loop 設定
            if (chkLoop != null) chkLoop.Checked = wf.LoopEnabled;
            if (wf.LoopEnabled)
            {
                if (chkLoopInfinite != null) chkLoopInfinite.Checked = wf.LoopCount == null;
                if (wf.LoopCount != null && numLoopCount != null) numLoopCount.Value = Math.Min(numLoopCount.Maximum, Math.Max(numLoopCount.Minimum, wf.LoopCount.Value));
                if (numLoopInterval != null) numLoopInterval.Value = Math.Min(numLoopInterval.Maximum, Math.Max(numLoopInterval.Minimum, (decimal)wf.LoopIntervalMs / 1000m));
            }
            btnReplay.Enabled = liveWorkflowSteps.Count > 0;
            btnClearRecord.Enabled = liveWorkflowSteps.Count > 0;
            UpdateStatus($"[SavedWF] 已載入流程: {name} (步驟 {liveWorkflowSteps.Count})");
        }

        private void btnSaveCurrent_Click(object? sender, EventArgs e)
        {
            if (liveWorkflowSteps == null || liveWorkflowSteps.Count == 0)
            {
                MessageBox.Show("目前沒有流程步驟可儲存", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string name = PromptForWorkflowName();
            if (string.IsNullOrWhiteSpace(name)) return;
            var wf = new Workflow
            {
                Name = name.Trim(),
                Steps = new List<AutoPressApp.Steps.Step>(liveWorkflowSteps),
                LoopEnabled = chkLoop != null && chkLoop.Checked,
                LoopCount = (chkLoop != null && chkLoop.Checked && chkLoopInfinite != null && chkLoopInfinite.Checked) ? (int?)null : (chkLoop != null && chkLoop.Checked ? (int?)numLoopCount.Value : null),
                LoopIntervalMs = (chkLoop != null && chkLoop.Checked && numLoopInterval != null) ? (int)(numLoopInterval.Value * 1000m) : 1000
            };
            try
            {
                bool overwrite = false;
                var existing = WorkflowStorage.Load(wf.Name);
                if (existing != null)
                {
                    if (MessageBox.Show("同名流程已存在，是否覆蓋?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                        return;
                    overwrite = true;
                }
                WorkflowStorage.Save(wf);
                LoadSavedWorkflowList();
                // 選回剛儲存的
                if (cmbSavedWorkflows != null)
                {
                    int idx = cmbSavedWorkflows.Items.IndexOf(wf.Name);
                    if (idx >= 0) cmbSavedWorkflows.SelectedIndex = idx;
                }
                UpdateStatus($"[SavedWF] {(overwrite ? "覆蓋" : "新增")}已儲存: {wf.Name}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("儲存失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnDeleteSaved_Click(object? sender, EventArgs e)
        {
            if (cmbSavedWorkflows == null || cmbSavedWorkflows.SelectedItem == null) return;
            string name = cmbSavedWorkflows.SelectedItem.ToString()!;
            if (MessageBox.Show($"刪除已儲存流程 '{name}'?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;
            if (WorkflowStorage.Delete(name))
            {
                UpdateStatus($"[SavedWF] 已刪除: {name}");
                LoadSavedWorkflowList();
            }
            else
            {
                MessageBox.Show("刪除失敗或檔案不存在", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string PromptForWorkflowName()
        {
            using var input = new InputBoxDialog("輸入流程名稱", "請輸入要儲存的流程名稱:");
            return input.ShowDialog(this) == DialogResult.OK ? input.ResultText : string.Empty;
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            LoadRunningApplications();
            UpdateStatus("應用程式清單已重新整理");
        }

        private void btnRecord_Click(object sender, EventArgs e)
        {
            // 改為錄製 KeySequenceStep 流程
            if (recorder == null)
            {
                recorder = new RecorderService();
                recorder.OnLog += UpdateStatus;
                recorder.StepCaptured += Recorder_StepCaptured;
                recorder.OnStopped += () =>
                {
                    // ESC 或其他方式停止錄製時自動結束 UI 狀態
                    if (System.Threading.Interlocked.Exchange(ref recordingFinalizeState, 1) == 0)
                    {
                        if (InvokeRequired)
                        {
                            Invoke(new Action(() => FinalizeWorkflowRecording(viaEsc: true, interactive: false, fromOnStopped:true)));
                        }
                        else
                        {
                            FinalizeWorkflowRecording(viaEsc: true, interactive: false, fromOnStopped:true);
                        }
                    }
                    if (btnRecord != null)
                    {
                        btnRecord.Text = "開始記錄";
                        btnRecord.BackColor = SystemColors.Control;
                        btnRecord.ForeColor = SystemColors.ControlText;
                    }
                    SetMode(RunMode.Idle);
                };
                liveWorkflowSteps = new List<AutoPressApp.Steps.Step>();
                if (lstRecordedKeys != null) lstRecordedKeys.Items.Clear();
                System.Threading.Interlocked.Exchange(ref recordingFinalizeState, 0);
                recorder.Start();
                UpdateStatus("[Recorder] 開始錄製 (再次點擊停止並儲存)");
                SetMode(RunMode.RecordingWorkflow);
                btnRecord.Text = "停止記錄";
                btnRecord.BackColor = Color.Red;
                btnRecord.ForeColor = Color.White;
            }
            else
            {
                FinalizeWorkflowRecording(viaEsc: false, interactive: true);
                btnRecord.Text = "開始記錄";
                btnRecord.BackColor = SystemColors.Control;
                btnRecord.ForeColor = SystemColors.ControlText;
                SetMode(RunMode.Idle);
            }
        }

        private void btnReplay_Click(object sender, EventArgs e)
        {
            UpdateStatus("[BUTTON] btnReplay_Click 被點擊");
            // (已移除舊 isRecording 流程切換)
            if (liveWorkflowSteps == null || liveWorkflowSteps.Count == 0)
            {
                UpdateStatus("[BUTTON] 尚無錄製流程可回放");
                MessageBox.Show("尚無錄製流程可回放！請先錄製流程。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var wf = new Workflow { Name = "Live Replay", Steps = new List<AutoPressApp.Steps.Step>(liveWorkflowSteps) };
            _ = RunWorkflowPreviewAsync(wf);
        }

        private void btnClearRecord_Click(object sender, EventArgs e)
        {
            liveWorkflowSteps?.Clear();
            if (lstRecordedKeys != null) lstRecordedKeys.Items.Clear();
            if (tvSteps != null) tvSteps.Nodes.Clear();
            btnReplay.Enabled = false;
            btnClearRecord.Enabled = false;
            UpdateStatus("流程已清除");
        }

        private void Form1_FormClosing(object? sender, FormClosingEventArgs e)
        {
            if (keyboardHook != IntPtr.Zero)
            {
                UnhookWindowsHookEx(keyboardHook);
            }
        }

    // 單鍵模式間隔事件已移除

        // Legacy export/import methods replaced with Workflow JSON

    private void btnExport_Click(object sender, EventArgs e)
    {
        if (liveWorkflowSteps == null || liveWorkflowSteps.Count == 0)
        {
            MessageBox.Show("目前沒有流程可匯出", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        using var sfd = new SaveFileDialog { Filter = "Workflow JSON (*.json)|*.json", FileName = "workflow.json" };
        if (sfd.ShowDialog() == DialogResult.OK)
        {
            try
            {
                var wf = new Workflow { Name = "Exported Workflow", Steps = new List<AutoPressApp.Steps.Step>(liveWorkflowSteps) };
                var json = WorkflowRunner.SaveToJson(wf);
                System.IO.File.WriteAllText(sfd.FileName, json, Encoding.UTF8);
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
        using var ofd = new OpenFileDialog { Filter = "Workflow JSON (*.json)|*.json" };
        if (ofd.ShowDialog() == DialogResult.OK)
        {
            try
            {
                var json = System.IO.File.ReadAllText(ofd.FileName, Encoding.UTF8);
                var wf = WorkflowRunner.LoadFromJson(json);
                liveWorkflowSteps = new List<AutoPressApp.Steps.Step>(wf.Steps);
                RebuildTreeView();
                btnReplay.Enabled = liveWorkflowSteps.Count > 0;
                btnClearRecord.Enabled = liveWorkflowSteps.Count > 0;
                UpdateStatus($"已匯入 {liveWorkflowSteps.Count} 個流程步驟");
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
                delayMultiplier = v; // 播放速度倍率：2.0x表示2倍速，0.5x表示半速
                UpdateStatus($"速度倍率設定: {txt} (delayMultiplier={delayMultiplier})");
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

        private async void RunWorkflowViaPicker()
        {
            using var ofd = new OpenFileDialog { Filter = "Workflow JSON (*.json)|*.json", FileName = "sample-workflow.json" };
            if (ofd.ShowDialog() != DialogResult.OK) return;
            try
            {
                string json = System.IO.File.ReadAllText(ofd.FileName);
                var wf = WorkflowRunner.LoadFromJson(json);
                var log = new LogService();
                log.OnLog += (m) => UpdateStatus(m);
                var runner = new WorkflowRunner(log, delayMultiplier);
                // 設定派送模式
                runner.DispatchMode = (cmbDispatchMode != null && cmbDispatchMode.SelectedIndex == 1)
                    ? Steps.InputDispatchMode.BackgroundPostMessage
                    : Steps.InputDispatchMode.ForegroundSendInput;
                runner.TargetWindowHandle = ResolveSelectedWindowHandle();
                workflowCts?.Cancel();
                workflowCts = new CancellationTokenSource();
                await runner.RunAsync(wf, workflowCts.Token);
            }
            catch (Exception ex)
            {
                MessageBox.Show("執行流程失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
    }

        private IntPtr ResolveSelectedWindowHandle()
        {
            try
            {
                if (cmbApplications == null || cmbApplications.SelectedIndex <= 0)
                {
                    return GetForegroundWindow(); // (當前焦點視窗)
                }
                // 項目格式: "WindowTitle (ProcessName|0xHANDLE)" 或其他自定義，這裡嘗試解析最後的 HANDLE
                var text = cmbApplications.SelectedItem?.ToString();
                if (string.IsNullOrEmpty(text)) return IntPtr.Zero;
                int idx = text.LastIndexOf("0x", StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                {
                    var hex = text.Substring(idx + 2).TrimEnd(')');
                    if (IntPtr.Size == 8 && ulong.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out var h64))
                        return (IntPtr)(long)h64;
                    if (IntPtr.Size == 4 && uint.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out var h32))
                        return (IntPtr)(int)h32;
                }
            }
            catch { }
            return IntPtr.Zero;
        }

        private void ToggleRecordWorkflow()
        {
            try
            {
                if (currentMode == RunMode.PlayingWorkflow)
                {
                    UpdateStatus("[Recorder] 播放中無法開始錄製");
                    return;
                }
                if (recorder == null)
                {
                    recorder = new RecorderService();
                    recorder.OnLog += UpdateStatus;
                    recorder.StepCaptured += Recorder_StepCaptured;
                    liveWorkflowSteps = new List<AutoPressApp.Steps.Step>();
                    if (lstRecordedKeys != null) lstRecordedKeys.Items.Clear();
                    System.Threading.Interlocked.Exchange(ref recordingFinalizeState, 0);
                    recorder.Start();
                    UpdateStatus("[Recorder] 開始錄製 (Ctrl+Shift+R 再次停止並儲存)");
                    SetMode(RunMode.RecordingWorkflow);
                }
                else
                {
                    FinalizeWorkflowRecording(viaEsc: false, interactive: true);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus("[Recorder] 錯誤: " + ex.Message);
            }
        }

        private void FinalizeWorkflowRecording(bool viaEsc, bool interactive, bool fromOnStopped = false)
        {
            try
            {
                // 防止重入 (若尚未由 OnStopped 設置，則設定 flag)
                if (System.Threading.Interlocked.CompareExchange(ref recordingFinalizeState, 1, 0) != 0 && !fromOnStopped)
                {
                    return; // 已 finalize
                }
                if (recorder == null)
                {
                    return;
                }
                var wf = recorder.StopToWorkflow("Recorded Workflow");
                try { recorder.OnLog -= UpdateStatus; } catch { }
                try { recorder.StepCaptured -= Recorder_StepCaptured; } catch { }
                try { recorder.Dispose(); } catch { }
                recorder = null;
                liveWorkflowSteps = new List<AutoPressApp.Steps.Step>(wf.Steps);
                RebuildTreeView();
                SetMode(RunMode.Idle);
                UpdateStatus(viaEsc ? "[Recorder] 錄製完成 (ESC) 已保留步驟" : "[Recorder] 錄製完成");
                // 取消原本的互動式預覽與自動儲存對話框，避免每次錄製結束彈出視窗。
                // 使用者可手動按下『回放記錄』或『匯出』來預覽 / 儲存。
                if (interactive)
                {
                    UpdateStatus("[Recorder] 已取消自動預覽/儲存 (可手動使用 回放記錄 / 匯出)");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus("[Recorder] Finalize 錯誤: " + ex.Message);
            }
        }

        private void Recorder_StepCaptured(AutoPressApp.Steps.Step step, string summary)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => Recorder_StepCaptured(step, summary)));
                return;
            }
            liveWorkflowSteps?.Add(step);
            AppendStepToUI(step, liveWorkflowSteps!.Count - 1, summary);
        }

        private async Task RunWorkflowPreviewAsync(Workflow wf)
        {
            var log = new LogService();
            log.OnLog += m => UpdateStatus(m);
            var runner = new WorkflowRunner(log, delayMultiplier);
            DateTime startTime = DateTime.Now;
            void EnsureClockBaseline()
            {
                if (lblClock == null) return;
                if (!lblClock.IsHandleCreated) return;
                var lines = lblClock.Text.Split('\n');
                if (lines.Length < 5)
                {
                    lblClock.Text = "Now --:--:--\nElapsed 00:00:00\nLoop -/-\nStep -/-\nNext --";
                }
            }
            void UpdateClock(Action<string[]> mutator)
            {
                if (lblClock == null) return;
                try
                {
                    lblClock.Invoke(new Action(() =>
                    {
                        EnsureClockBaseline();
                        var arr = lblClock.Text.Split('\n');
                        if (arr.Length < 5) return;
                        mutator(arr);
                        lblClock.Text = string.Join("\n", arr);
                    }));
                }
                catch { }
            }
            var clockUpdaterCts = new CancellationTokenSource();
            _ = Task.Run(async () =>
            {
                while (!clockUpdaterCts.IsCancellationRequested)
                {
                    if (lblClock != null)
                    {
                        var now = DateTime.Now;
                        var elapsed = now - startTime;
                        UpdateClock(arr =>
                        {
                            arr[0] = $"Now {now:HH:mm:ss}";
                            arr[1] = $"Elapsed {elapsed:hh\\:mm\\:ss}";
                        });
                    }
                    try { await Task.Delay(1000, clockUpdaterCts.Token); } catch { break; }
                }
            }, clockUpdaterCts.Token);
            // Apply loop settings from UI if enabled
            try
            {
                if (chkLoop != null && chkLoop.Checked)
                {
                    wf.LoopEnabled = true;
                    if (chkLoopInfinite != null && chkLoopInfinite.Checked)
                    {
                        wf.LoopCount = null; // interpreted as infinite loops (runner uses int.MaxValue)
                    }
                    else if (numLoopCount != null)
                    {
                        wf.LoopCount = (int)numLoopCount.Value;
                    }
                    if (numLoopInterval != null)
                    {
                        // 使用者以秒輸入，轉成毫秒
                        wf.LoopIntervalMs = (int)(numLoopInterval.Value * 1000m);
                    }
                }
                else
                {
                    wf.LoopEnabled = false;
                }
            }
            catch { }
            runner.OnStepExecuting += (idx, step) =>
            {
                if (lstRecordedKeys != null && idx >= 0 && idx < lstRecordedKeys.Items.Count)
                {
                    lstRecordedKeys.SelectedIndex = idx;
                    lstRecordedKeys.Refresh();
                }
                UpdateStatus($"[Preview] 執行步驟 {idx + 1}/{wf.Steps.Count}: {step.GetType().Name}");
                UpdateClock(arr => arr[3] = $"Step {idx + 1}/{wf.Steps.Count}");
            };
            runner.OnLoopStarting += (curr, total, infinite) =>
            {
                UpdateClock(arr =>
                {
                    arr[2] = $"Loop {curr}/{(infinite?"∞":total.ToString())}";
                    arr[4] = "Next --";
                });
            };
            runner.OnIntervalTick += (remainingMs) =>
            {
                int sec = remainingMs / 1000;
                int cs = (remainingMs % 1000) / 10;
                UpdateClock(arr => arr[4] = remainingMs > 0 ? $"Next {sec:D2}.{cs:D2}s" : "Next RUN");
            };
            workflowCts?.Cancel();
            workflowCts = new CancellationTokenSource();
            try
            {
                UpdateStatus("[Preview] 執行錄製流程...");
                SetMode(RunMode.PlayingWorkflow);
                isRunning = true;
                UpdateStatus($"[Preview] 步驟數: {wf.Steps.Count}");
                await runner.RunAsync(wf, workflowCts.Token);
                UpdateStatus("[Preview] 完成");
            }
            catch (Exception ex)
            {
                UpdateStatus("[Preview] 失敗: " + ex.Message);
            }
            finally
            {
                isRunning = false;
                if (currentMode == RunMode.PlayingWorkflow)
                    SetMode(RunMode.Idle);
                ReindexTreeNodes();
                try { clockUpdaterCts.Cancel(); } catch { }
            }
        }

    // btnWorkflowMenu_Click 已移除

        private void chkLoop_CheckedChanged(object? sender, EventArgs e)
        {
            bool en = chkLoop.Checked;
            if (numLoopCount != null) { numLoopCount.Enabled = en; }
            if (numLoopInterval != null) { numLoopInterval.Enabled = en; }
            if (lblLoopCount != null) { lblLoopCount.Enabled = en; }
            if (lblLoopInterval != null) { lblLoopInterval.Enabled = en; }
            UpdateStatus(en ? "[Loop] 循環已啟用" : "[Loop] 循環已停用");
        }

        private void chkLoopInfinite_CheckedChanged(object? sender, EventArgs e)
        {
            if (chkLoopInfinite == null) return;
            bool inf = chkLoopInfinite.Checked;
            if (numLoopCount != null) numLoopCount.Enabled = chkLoop != null && chkLoop.Checked && !inf;
            if (lblLoopCount != null) lblLoopCount.Enabled = chkLoop != null && chkLoop.Checked && !inf;
            UpdateStatus(inf ? "[Loop] 無限循環" : "[Loop] 次數循環");
        }
        // --- TreeView 支援 ---
        private void AppendStepToUI(AutoPressApp.Steps.Step step, int index, string summary)
        {
            if (tvSteps == null || !tvSteps.Visible)
            {
                if (lstRecordedKeys != null)
                {
                    lstRecordedKeys.Items.Add(summary);
                    lstRecordedKeys.TopIndex = lstRecordedKeys.Items.Count - 1;
                }
                return;
            }
            if (index < 0) return;
            var node = BuildStepNode(step, index);
            tvSteps.Nodes.Add(node);
            tvSteps.SelectedNode = node;
            node.EnsureVisible();
        }

        private TreeNode BuildStepNode(AutoPressApp.Steps.Step step, int index)
        {
            string caption = step switch
            {
                AutoPressApp.Steps.DelayStep d => $"[{index}] Delay {d.Ms}ms",
                AutoPressApp.Steps.KeyComboStep k => $"[{index}] KeyCombo {string.Join('+', k.Keys)}",
                AutoPressApp.Steps.KeySequenceStep ks => $"[{index}] 按鍵序列 {ks.Events.Count} 事件",
                AutoPressApp.Steps.MouseClickStep m => $"[{index}] Click {m.Button} ({m.X},{m.Y})",
                AutoPressApp.Steps.FocusWindowStep f => $"[{index}] Focus '{Truncate(f.TitleContains,30)}'",
                AutoPressApp.Steps.LogStep lg => $"[{index}] Log {Truncate(lg.Message,30)}",
                _ => $"[{index}] {step.GetType().Name}"
            };
            var root = new TreeNode(caption) { Tag = index };
            if (step is AutoPressApp.Steps.KeySequenceStep kseq)
            {
                int i = 0;
                foreach (var ev in kseq.Events)
                {
                    root.Nodes.Add(new TreeNode($"{i++:D3}: {(ev.Down?"Down":"Up  ")} {ev.Key} (+{ev.DelayMsBefore}ms)"));
                }
            }
            return root;
        }

        private void RebuildTreeView()
        {
            if (tvSteps == null) return;
            tvSteps.BeginUpdate();
            tvSteps.Nodes.Clear();
            if (liveWorkflowSteps != null)
            {
                for (int i = 0; i < liveWorkflowSteps.Count; i++)
                {
                    tvSteps.Nodes.Add(BuildStepNode(liveWorkflowSteps[i], i));
                }
            }
            tvSteps.EndUpdate();
        }

        private void ReindexTreeNodes()
        {
            if (tvSteps == null || liveWorkflowSteps == null) return;
            for (int i = 0; i < tvSteps.Nodes.Count && i < liveWorkflowSteps.Count; i++)
            {
                var n = tvSteps.Nodes[i];
                n.Text = BuildStepNode(liveWorkflowSteps[i], i).Text;
                n.Tag = i;
            }
        }

        private static string Truncate(string? s, int max)
        {
            if (string.IsNullOrEmpty(s)) return string.Empty;
            return s.Length <= max ? s : s.Substring(0, max - 3) + "...";
        }

        // ---- TreeView 編輯支援 ----
        private ContextMenuStrip? tvContextMenu;
        private void EnsureTreeViewContext()
        {
            if (tvSteps == null) return;
            if (tvContextMenu != null) return;
            tvContextMenu = new ContextMenuStrip();
            tvContextMenu.Items.Add("編輯", null, (_, __) => EditSelectedStep());
            tvContextMenu.Items.Add("上移", null, (_, __) => MoveSelectedStep(-1));
            tvContextMenu.Items.Add("下移", null, (_, __) => MoveSelectedStep(1));
            tvContextMenu.Items.Add("刪除", null, (_, __) => DeleteSelectedStep());
            // 子事件操作
            tvContextMenu.Items.Add(new ToolStripSeparator());
            tvContextMenu.Items.Add("編輯子事件", null, (_, __) => EditKeySequenceEvent());
            tvContextMenu.Items.Add("刪除子事件", null, (_, __) => DeleteKeySequenceEvent());
            tvContextMenu.Items.Add("子事件前插入", null, (_, __) => InsertKeySequenceEvent(before:true));
            tvContextMenu.Items.Add("子事件後插入", null, (_, __) => InsertKeySequenceEvent(before:false));
            tvContextMenu.Items.Add(new ToolStripSeparator());
            tvContextMenu.Items.Add("重新索引", null, (_, __) => ReindexTreeNodes());
            tvSteps.ContextMenuStrip = tvContextMenu;
            tvSteps.NodeMouseClick += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                {
                    tvSteps.SelectedNode = e.Node;
                }
            };
            tvSteps.DoubleClick += (s, e) => EditSelectedStep();
        }

        private int? GetSelectedRootIndex()
        {
            if (tvSteps == null || liveWorkflowSteps == null) return null;
            var node = tvSteps.SelectedNode;
            if (node == null) return null;
            // 允許子節點選擇時，取其 Root
            if (node.Parent != null) node = node.Parent;
            if (node.Tag is int idx && idx >= 0 && idx < liveWorkflowSteps.Count) return idx;
            // 如果 Tag 不在，根據其在根集合位置推斷
            int pos = node.Level == 0 ? node.Index : node.Parent!.Index;
            if (pos >= 0 && pos < liveWorkflowSteps.Count) return pos;
            return null;
        }

        private void EditSelectedStep()
        {
            var idx = GetSelectedRootIndex();
            if (idx == null || liveWorkflowSteps == null) return;
            var step = liveWorkflowSteps[idx.Value];
            using var dlg = new StepEditDialog(step);
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                ReindexTreeNodes();
                UpdateStatus($"[Edit] 已更新步驟 {idx.Value}");
            }
        }

        private void MoveSelectedStep(int delta)
        {
            var idx = GetSelectedRootIndex();
            if (idx == null || liveWorkflowSteps == null) return;
            int newIndex = idx.Value + delta;
            if (newIndex < 0 || newIndex >= liveWorkflowSteps.Count) return;
            var tmp = liveWorkflowSteps[idx.Value];
            liveWorkflowSteps.RemoveAt(idx.Value);
            liveWorkflowSteps.Insert(newIndex, tmp);
            RebuildTreeView();
            if (tvSteps != null && newIndex < tvSteps.Nodes.Count)
            {
                tvSteps.SelectedNode = tvSteps.Nodes[newIndex];
                tvSteps.Nodes[newIndex].EnsureVisible();
            }
            UpdateStatus($"[Edit] 步驟已移動到 {newIndex}");
        }

        private void DeleteSelectedStep()
        {
            var idx = GetSelectedRootIndex();
            if (idx == null || liveWorkflowSteps == null) return;
            if (MessageBox.Show($"確定刪除步驟 {idx}?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                liveWorkflowSteps.RemoveAt(idx.Value);
                RebuildTreeView();
                btnReplay.Enabled = liveWorkflowSteps.Count > 0;
                btnClearRecord.Enabled = liveWorkflowSteps.Count > 0;
                UpdateStatus($"[Edit] 已刪除步驟 {idx}");
            }
        }

        // 在表單載入完成後設定 TreeView 右鍵
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            EnsureTreeViewContext();
        }

        // ---- KeySequence 子事件編輯 ----
    private (AutoPressApp.Steps.KeySequenceStep step, int eventIndex)? GetSelectedKeySequenceEvent()
        {
            if (tvSteps == null || liveWorkflowSteps == null) return null;
            var node = tvSteps.SelectedNode;
            if (node == null) return null;
            if (node.Parent == null) return null; // 根節點不是子事件
            // 父節點 index 對應步驟
            int stepIdx = node.Parent.Index;
            if (stepIdx < 0 || stepIdx >= liveWorkflowSteps.Count) return null;
            if (liveWorkflowSteps[stepIdx] is AutoPressApp.Steps.KeySequenceStep ks)
            {
                int childIdx = node.Index;
                if (childIdx >= 0 && childIdx < ks.Events.Count)
                    return (ks, childIdx);
            }
            return null;
        }

        private void EditKeySequenceEvent()
        {
            var info = GetSelectedKeySequenceEvent();
            if (info == null) return;
            var (ks, idx) = info.Value;
            var item = ks.Events[idx];
            using var dlg = new KeyEventEditDialog(item, idx);
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                RebuildTreeView();
                UpdateStatus($"[Edit] 已更新子事件 {idx}");
            }
        }

        private void DeleteKeySequenceEvent()
        {
            var info = GetSelectedKeySequenceEvent();
            if (info == null) return;
            var (ks, idx) = info.Value;
            if (MessageBox.Show($"刪除子事件 {idx}?", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                ks.Events.RemoveAt(idx);
                RebuildTreeView();
                UpdateStatus($"[Edit] 已刪除子事件 {idx}");
            }
        }

        private void InsertKeySequenceEvent(bool before)
        {
            var info = GetSelectedKeySequenceEvent();
            if (info == null) return;
            var (ks, idx) = info.Value;
            int insertIdx = before ? idx : idx + 1;
            var newItem = new AutoPressApp.Steps.KeyEventItem { Key = "A", Down = true, DelayMsBefore = 0 };
            ks.Events.Insert(insertIdx, newItem);
            using var dlg = new KeyEventEditDialog(newItem, insertIdx);
            if (dlg.ShowDialog(this) == DialogResult.OK)
            {
                RebuildTreeView();
                UpdateStatus($"[Edit] 已插入子事件 {insertIdx}");
            }
            else
            {
                // 取消則移除
                ks.Events.Remove(newItem);
            }
        }
    }
}
