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
    private const int KEYEVENTF_SCANCODE = 0x0008;
        private const int SW_RESTORE = 9;

    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        
    private System.Windows.Forms.Timer replayTimer = new System.Windows.Forms.Timer();
    private bool isRunning = false;
        private bool isRecording = false;
        private bool isReplaying = false; // 用於判斷目前是否在回放
        private IntPtr targetWindowHandle = IntPtr.Zero;
        private string targetWindowTitle = "";
        
        // 按鍵記錄相關
    // 移除 Legacy 鍵盤序列相關欄位
    private IntPtr keyboardHook = IntPtr.Zero;
    private LowLevelKeyboardProc hookProc;
    private double delayMultiplier = 1.0;
    private bool LoopPlayback => chkLoop != null && chkLoop.Checked;

        // 工作流程相關
        private CancellationTokenSource? workflowCts;
        private RecorderService? recorder;
    private List<AutoPressApp.Steps.Step>? liveWorkflowSteps;

        private ContextMenuStrip? workflowMenu;
        // 模式狀態 (統一顯示 Recording / Playing)
        private enum RunMode { Idle, RecordingWorkflow, RecordingLegacy, PlayingWorkflow, PlayingLegacy }
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
                    RunMode.RecordingLegacy => "狀態: 錄製舊按鍵中 (ESC 停止)",
                    RunMode.PlayingWorkflow => "狀態: 回放流程中 (ESC 停止)",
                    RunMode.PlayingLegacy => "狀態: 回放舊按鍵中 (ESC 停止)",
                    _ => "狀態: ?"
                };
                if (lblMode != null) lblMode.Text = txt;
                Color back; Color fore;
                switch (mode)
                {
                    case RunMode.RecordingWorkflow:
                    case RunMode.RecordingLegacy:
                        back = Color.Red; fore = Color.White; break;
                    case RunMode.PlayingWorkflow:
                    case RunMode.PlayingLegacy:
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
                    if (mode == RunMode.PlayingWorkflow || mode == RunMode.PlayingLegacy)
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

        private void InitializeTimers()
        {
            // Legacy timer removed
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
                    ToggleRecordWorkflow();
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
                        FinalizeWorkflowRecording(viaEsc: true, interactive: false);
                        UpdateStatus("[Recorder] ESC 結束並保留流程");
                    }
                    else
                    {
                        StopAll();
                        UpdateStatus("已透過 ESC 停止所有執行");
                    }
                    return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
                }

                if (isRecording)
                {
                    // Legacy recording removed - now uses RecorderService for KeySequenceStep
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
        
        private void SendInputToTarget(Keys key)
        {
            UpdateStatus($"[LOG] 開始發送按鍵到系統: {key}");
            
            // 確保目標視窗是焦點視窗（如果有指定的話）
            if (targetWindowHandle != IntPtr.Zero)
            {
                bool setForeground = SetForegroundWindow(targetWindowHandle);
                UpdateStatus($"[LOG] 設置目標視窗為焦點: {setForeground}");
                
                // 等待視窗切換完成
                System.Threading.Thread.Sleep(100);
                
                // 確認前景視窗
                IntPtr currentForeground = GetForegroundWindow();
                bool isForeground = currentForeground == targetWindowHandle;
                UpdateStatus($"[LOG] 前景視窗確認: Expected={targetWindowHandle}, Actual={currentForeground}, Match={isForeground}");
            }
            
            // 使用全域 SendInput 發送到系統（不指定特定視窗）
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

        private bool TryClipboardMethod(Keys key)
        {
            try
            {
                // 將按鍵轉換為字符
                char character = GetCharFromKey(key);
                if (character == '\0' || char.IsControl(character))
                {
                    return false; // 只處理可打印字符
                }
                
                UpdateStatus($"[LOG] 嘗試剪貼板方法發送: '{character}'");
                
                // 保存當前剪貼板內容
                string originalClipboard = "";
                if (System.Windows.Forms.Clipboard.ContainsText())
                {
                    originalClipboard = System.Windows.Forms.Clipboard.GetText();
                }
                
                // 將字符放入剪貼板
                System.Windows.Forms.Clipboard.SetText(character.ToString());
                
                // 發送 Ctrl+V 到目標視窗
                const uint WM_KEYDOWN = 0x0100;
                const uint WM_KEYUP = 0x0101;
                const int VK_CONTROL = 0x11;
                const int VK_V = 0x56;
                
                // 按下 Ctrl
                SendMessage(targetWindowHandle, WM_KEYDOWN, (IntPtr)VK_CONTROL, IntPtr.Zero);
                System.Threading.Thread.Sleep(10);
                
                // 按下 V
                SendMessage(targetWindowHandle, WM_KEYDOWN, (IntPtr)VK_V, IntPtr.Zero);
                System.Threading.Thread.Sleep(10);
                
                // 釋放 V
                SendMessage(targetWindowHandle, WM_KEYUP, (IntPtr)VK_V, IntPtr.Zero);
                System.Threading.Thread.Sleep(10);
                
                // 釋放 Ctrl
                SendMessage(targetWindowHandle, WM_KEYUP, (IntPtr)VK_CONTROL, IntPtr.Zero);
                
                // 等待一下讓貼上完成
                System.Threading.Thread.Sleep(100);
                
                // 恢復原來的剪貼板內容
                if (!string.IsNullOrEmpty(originalClipboard))
                {
                    System.Windows.Forms.Clipboard.SetText(originalClipboard);
                }
                else
                {
                    System.Windows.Forms.Clipboard.Clear();
                }
                
                UpdateStatus($"[LOG] 剪貼板方法完成");
                return true;
            }
            catch (Exception ex)
            {
                UpdateStatus($"[LOG] 剪貼板方法失敗: {ex.Message}");
                return false;
            }
        }

        private bool TrySendCharMessage(Keys key)
        {
            if (targetWindowHandle == IntPtr.Zero)
                return false;
                
            // 將 Keys 轉換為字符
            char character = GetCharFromKey(key);
            if (character == '\0')
            {
                UpdateStatus($"[LOG] 無法轉換按鍵為字符: {key}");
                return false;
            }
            
            // 獲取虛擬鍵碼
            ushort virtualKey = GetVirtualKeyCode(key);
            
            UpdateStatus($"[LOG] 嘗試發送完整按鍵序列: '{character}' (VK=0x{virtualKey:X2})");
            
            // 發送完整的按鍵序列：KEYDOWN -> CHAR -> KEYUP
            const uint WM_KEYDOWN = 0x0100;
            const uint WM_CHAR = 0x0102;
            const uint WM_KEYUP = 0x0101;
            
            // 1. 發送 WM_KEYDOWN
            IntPtr keyDownResult = SendMessage(targetWindowHandle, WM_KEYDOWN, (IntPtr)virtualKey, IntPtr.Zero);
            UpdateStatus($"[LOG] SendMessage WM_KEYDOWN 結果: {keyDownResult}");
            
            // 短暫延遲
            System.Threading.Thread.Sleep(10);
            
            // 2. 發送 WM_CHAR (只對可打印字符)
            if (char.IsControl(character) == false)
            {
                IntPtr charResult = SendMessage(targetWindowHandle, WM_CHAR, (IntPtr)character, IntPtr.Zero);
                UpdateStatus($"[LOG] SendMessage WM_CHAR 結果: {charResult}");
            }
            
            // 短暫延遲
            System.Threading.Thread.Sleep(10);
            
            // 3. 發送 WM_KEYUP
            IntPtr keyUpResult = SendMessage(targetWindowHandle, WM_KEYUP, (IntPtr)virtualKey, IntPtr.Zero);
            UpdateStatus($"[LOG] SendMessage WM_KEYUP 結果: {keyUpResult}");
            
            // 檢查是否至少有一個消息被處理
            bool anySuccess = keyDownResult != IntPtr.Zero || keyUpResult != IntPtr.Zero;
            UpdateStatus($"[LOG] 按鍵序列發送完成，成功: {anySuccess}");
            
            // 只有當消息確實被處理時才返回 true
            return anySuccess;
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

        // 測試SendInput功能
        private void TestSendInput()
        {
            UpdateStatus("[TEST] 測試 SendInput 功能 - 將發送字母 'A'");
            
            var inputs = new INPUT[]
            {
                new INPUT
                {
                    type = 1,
                    U = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = 0x41, // A key
                            wScan = 0,
                            dwFlags = 0,
                            time = 0,
                            dwExtraInfo = UIntPtr.Zero
                        }
                    }
                },
                new INPUT
                {
                    type = 1,
                    U = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = 0x41, // A key
                            wScan = 0,
                            dwFlags = KEYEVENTF_KEYUP,
                            time = 0,
                            dwExtraInfo = UIntPtr.Zero
                        }
                    }
                }
            };
            
            uint result = SendInput(2, inputs, Marshal.SizeOf<INPUT>());
            UpdateStatus($"[TEST] SendInput 測試結果: {result}/2");
            
            if (result == 0)
            {
                int error = Marshal.GetLastWin32Error();
                UpdateStatus($"[TEST] SendInput 失敗，錯誤: {error}");
            }
        }

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
            if (currentMode == RunMode.PlayingWorkflow || currentMode == RunMode.PlayingLegacy || isRunning)
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
            if (recorder != null || isRecording)
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
            // fallback legacy
            BeginLegacyKeySequenceIfPossible();
        }

        private void BeginLegacyKeySequenceIfPossible()
        {
            // Legacy method replaced - now uses liveWorkflowSteps
            UpdateStatus("[Start] 尚未錄製任何流程步驟，請先錄製流程。");
            MessageBox.Show("尚未錄製任何流程步驟。請先錄製流程。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            if (isRecording)
            {
                isRecording = false;
                if (btnRecord != null)
                {
                    btnRecord.Text = "開始記錄";
                    btnRecord.BackColor = SystemColors.Control;
                    btnRecord.ForeColor = SystemColors.ControlText;
                }
            }
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
                liveWorkflowSteps = new List<AutoPressApp.Steps.Step>();
                if (lstRecordedKeys != null) lstRecordedKeys.Items.Clear();
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
            if (isRecording)
            {
                UpdateStatus("[BUTTON] 停止記錄模式");
                isRecording = false;
                btnRecord.Text = "開始記錄";
                btnRecord.BackColor = SystemColors.Control;
                btnRecord.ForeColor = SystemColors.ControlText;
            }
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
                if (lstRecordedKeys != null)
                {
                    lstRecordedKeys.Items.Clear();
                    foreach (var s in liveWorkflowSteps)
                    {
                        lstRecordedKeys.Items.Add(s switch
                        {
                            AutoPressApp.Steps.DelayStep d => $"Delay {d.Ms}ms",
                            AutoPressApp.Steps.KeyComboStep k => $"KeyCombo {string.Join('+', k.Keys)}",
                            AutoPressApp.Steps.KeySequenceStep ks => $"KeySeq {ks.Events.Count} ev",
                            AutoPressApp.Steps.MouseClickStep m => $"Click {m.Button} ({m.X},{m.Y})",
                            AutoPressApp.Steps.FocusWindowStep f => $"Focus '{f.TitleContains}'",
                            AutoPressApp.Steps.LogStep lg => $"Log \"{lg.Message}\"",
                            _ => s.GetType().Name
                        });
                    }
                }
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
                var runner = new WorkflowRunner(log);
                workflowCts?.Cancel();
                workflowCts = new CancellationTokenSource();
                await runner.RunAsync(wf, workflowCts.Token);
            }
            catch (Exception ex)
            {
                MessageBox.Show("執行流程失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ToggleRecordWorkflow()
        {
            try
            {
                if (recorder == null)
                {
                    recorder = new RecorderService();
                    recorder.OnLog += UpdateStatus;
                    recorder.StepCaptured += Recorder_StepCaptured;
                    liveWorkflowSteps = new List<AutoPressApp.Steps.Step>();
                    if (lstRecordedKeys != null) lstRecordedKeys.Items.Clear();
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

        private void FinalizeWorkflowRecording(bool viaEsc, bool interactive)
        {
            try
            {
                if (recorder == null) return;
                var wf = recorder.StopToWorkflow("Recorded Workflow");
                recorder.OnLog -= UpdateStatus;
                recorder.StepCaptured -= Recorder_StepCaptured;
                recorder.Dispose();
                recorder = null;
                liveWorkflowSteps = new List<AutoPressApp.Steps.Step>(wf.Steps);
                // Refresh list to ensure last delay is shown
                if (lstRecordedKeys != null)
                {
                    lstRecordedKeys.Items.Clear();
                    foreach (var s in liveWorkflowSteps)
                    {
                        lstRecordedKeys.Items.Add(s switch
                        {
                            AutoPressApp.Steps.DelayStep d => $"Delay {d.Ms}ms",
                            AutoPressApp.Steps.KeyComboStep k => $"KeyCombo {string.Join('+', k.Keys)}",
                            AutoPressApp.Steps.KeySequenceStep ks => $"KeySeq {ks.Events.Count} ev",
                            AutoPressApp.Steps.MouseClickStep m => $"Click {m.Button} ({m.X},{m.Y})",
                            AutoPressApp.Steps.FocusWindowStep f => $"Focus '{f.TitleContains}'",
                            AutoPressApp.Steps.LogStep lg => $"Log \"{lg.Message}\"",
                            _ => s.GetType().Name
                        });
                    }
                }
                SetMode(RunMode.Idle);
                UpdateStatus(viaEsc ? "[Recorder] 錄製完成 (ESC) 已保留步驟" : "[Recorder] 錄製完成");
                if (interactive)
                {
                    var preview = MessageBox.Show("要立即回放一次以預覽錄製效果嗎?", "錄製完成", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (preview == DialogResult.Yes)
                    {
                        _ = RunWorkflowPreviewAsync(wf);
                    }
                    using var sfd = new SaveFileDialog { Filter = "Workflow JSON (*.json)|*.json", FileName = $"workflow-{DateTime.Now:yyyyMMdd-HHmmss}.json" };
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        var json = WorkflowRunner.SaveToJson(wf);
                        System.IO.File.WriteAllText(sfd.FileName, json, Encoding.UTF8);
                        UpdateStatus($"[Recorder] 已儲存: {sfd.FileName}");
                    }
                    else
                    {
                        UpdateStatus("[Recorder] 已取消儲存");
                    }
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
            if (lstRecordedKeys != null)
            {
                lstRecordedKeys.Items.Add(step is AutoPressApp.Steps.KeySequenceStep seq
                    ? $"KeySeq {seq.Events.Count} ev"
                    : summary);
                lstRecordedKeys.TopIndex = lstRecordedKeys.Items.Count - 1; // auto-scroll
            }
        }

        private async Task RunWorkflowPreviewAsync(Workflow wf)
        {
            var log = new LogService();
            log.OnLog += m => UpdateStatus(m);
            var runner = new WorkflowRunner(log);
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
            }
        }

        private void btnWorkflowMenu_Click(object? sender, EventArgs e)
        {
            if (workflowMenu == null)
            {
                workflowMenu = new ContextMenuStrip();
                workflowMenu.Items.Add("錄製 (Ctrl+Shift+R)", null, (_, __) => ToggleRecordWorkflow());
                workflowMenu.Items.Add("執行 (選擇 Workflow JSON) (Ctrl+Shift+W)", null, (_, __) => RunWorkflowViaPicker());
                workflowMenu.Items.Add("重新預覽最近錄製", null, (_, __) =>
                {
                    if (liveWorkflowSteps != null && liveWorkflowSteps.Count > 0)
                    {
                        var wf = new Workflow { Name = "Live Preview", Steps = new List<AutoPressApp.Steps.Step>(liveWorkflowSteps) };
                        _ = RunWorkflowPreviewAsync(wf);
                    }
                    else
                    {
                        UpdateStatus("[Workflow] 尚無錄製內容可預覽");
                    }
                });
            }
            workflowMenu.Show(btnWorkflowMenu, new System.Drawing.Point(0, btnWorkflowMenu.Height));
        }
    }
}
