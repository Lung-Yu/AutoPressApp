using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices;
using AutoPressApp.Models;
using AutoPressApp.Steps;

namespace AutoPressApp.Services
{
    // Minimal workflow recorder: captures mouse click events and window focus changes.
    public class RecorderService : IDisposable
    {
    private const int WH_MOUSE_LL = 14;
    private const int WH_KEYBOARD_LL = 13;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_LBUTTONUP = 0x0202;
        private const int WM_RBUTTONDOWN = 0x0204;
        private const int WM_RBUTTONUP = 0x0205;
        private const int WM_MOUSEMOVE = 0x0200;
    private const int WM_KEYDOWN = 0x0100;
    private const int WM_KEYUP = 0x0101;

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
    private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")] private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);
    [DllImport("user32.dll")] private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll")] private static extern IntPtr GetModuleHandle(string lpModuleName);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

    private IntPtr _mouseHook = IntPtr.Zero;
    private IntPtr _keyboardHook = IntPtr.Zero;
        private LowLevelMouseProc? _mouseProc;
    private LowLevelKeyboardProc? _keyboardProc;

        private DateTime _lastEventTime;
        private IntPtr _lastWindow = IntPtr.Zero;
        private readonly List<Step> _steps = new List<Step>();

        private (string? Button, POINT Pt, DateTime Time)? _pendingDown;

        public event Action<string>? OnLog;
    public event Action<Step, string>? StepCaptured;
    public IReadOnlyList<Step> Steps => _steps.AsReadOnly();

        public void Start()
        {
            _steps.Clear();
            _pendingDown = null;
            _lastEventTime = DateTime.Now;
            _lastWindow = GetForegroundWindow();
            AddFocusWindowStepIfChanged(_lastWindow, true);

            _mouseProc = MouseHookCallback;
            using var cur = Process.GetCurrentProcess();
            using var mod = cur.MainModule!;
            _mouseHook = SetWindowsHookEx(WH_MOUSE_LL, _mouseProc, GetModuleHandle(mod.ModuleName ?? string.Empty), 0);
            _keyboardProc = KeyboardHookCallback;
            _keyboardHook = SetWindowsHookEx(WH_KEYBOARD_LL, _keyboardProc, GetModuleHandle(mod.ModuleName ?? string.Empty), 0);
            Log("[Recorder] Started (mouse+keyboard)");
        }

        public Workflow StopToWorkflow(string name = "Recorded Workflow")
        {
            // 確保序列提交
            FinalizePendingSequence();
            if (_mouseHook != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_mouseHook);
                _mouseHook = IntPtr.Zero;
            }
            if (_keyboardHook != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_keyboardHook);
                _keyboardHook = IntPtr.Zero;
            }
            Log("[Recorder] Stopped");
            return new Workflow
            {
                Name = name,
                Steps = new List<Step>(_steps),
                LoopEnabled = false
            };
        }

        private void AddDelayFrom(DateTime from)
        {
            var now = DateTime.Now;
            int ms = (int)Math.Max(0, (now - from).TotalMilliseconds);
            if (ms > 0)
            {
                var step = new DelayStep { Ms = ms };
                _steps.Add(step);
                var line = $"Delay {ms}ms";
                Log($"[Recorder] +{line}");
                StepCaptured?.Invoke(step, line);
            }
            _lastEventTime = now;
        }

        private void AddFocusWindowStepIfChanged(IntPtr hwnd, bool force = false)
        {
            if (hwnd == IntPtr.Zero) return;
            if (!force && hwnd == _lastWindow) return;

            string title = ReadWindowTitle(hwnd);
            string? procName = ReadProcessName(hwnd);
            var focusStep = new FocusWindowStep { TitleContains = string.IsNullOrWhiteSpace(title) ? null : title, ProcessName = procName, TimeoutMs = 5000 };
            _steps.Add(focusStep);
            _lastWindow = hwnd;
            var line = $"FocusWindow '{title}' proc='{procName}'";
            Log($"[Recorder] +{line}");
            StepCaptured?.Invoke(focusStep, line);
        }

        private string ReadWindowTitle(IntPtr hwnd)
        {
            var sb = new System.Text.StringBuilder(256);
            GetWindowText(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private string? ReadProcessName(IntPtr hwnd)
        {
            try
            {
                GetWindowThreadProcessId(hwnd, out var pid);
                if (pid == 0) return null;
                var p = Process.GetProcessById((int)pid);
                return p.ProcessName;
            }
            catch
            {
                return null;
            }
        }

        private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var msg = (int)wParam;
                var data = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);

                // window focus change detection
                var fg = GetForegroundWindow();
                AddFocusWindowStepIfChanged(fg);

                switch (msg)
                {
                    case WM_LBUTTONDOWN:
                        _pendingDown = ("left", data.pt, DateTime.Now);
                        break;
                    case WM_RBUTTONDOWN:
                        _pendingDown = ("right", data.pt, DateTime.Now);
                        break;
                    case WM_LBUTTONUP:
                        if (_pendingDown.HasValue && _pendingDown.Value.Button == "left")
                        {
                            AddDelayFrom(_lastEventTime);
                            var step = new MouseClickStep { X = data.pt.x, Y = data.pt.y, Button = "left", Mode = CoordMode.Screen };
                            _steps.Add(step);
                            var line = $"Click L ({data.pt.x},{data.pt.y})";
                            Log($"[Recorder] +{line}");
                            StepCaptured?.Invoke(step, line);
                            _pendingDown = null;
                        }
                        break;
                    case WM_RBUTTONUP:
                        if (_pendingDown.HasValue && _pendingDown.Value.Button == "right")
                        {
                            AddDelayFrom(_lastEventTime);
                            var step = new MouseClickStep { X = data.pt.x, Y = data.pt.y, Button = "right", Mode = CoordMode.Screen };
                            _steps.Add(step);
                            var line = $"Click R ({data.pt.x},{data.pt.y})";
                            Log($"[Recorder] +{line}");
                            StepCaptured?.Invoke(step, line);
                            _pendingDown = null;
                        }
                        break;
                    case WM_MOUSEMOVE:
                        // ignore for now (can be used for future path/composite steps)
                        break;
                }
            }
            return CallNextHookEx(_mouseHook, nCode, wParam, lParam);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        private bool _isCtrl, _isShift, _isAlt, _isWin;

        // 內建程式控制用快捷鍵 (不應被錄製進流程)
        private static readonly System.Collections.Generic.HashSet<string> ReservedCombos = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Esc",              // 停止
            "Ctrl+Shift+R",     // 錄製開始/停止
            "Ctrl+Shift+S",     // StartPreferred
            "Ctrl+Shift+W",     // 載入 JSON
            "Ctrl+Shift+P",     // 撥放流程
            "Ctrl+Shift+X"      // StopAll
        };

    private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                var data = Marshal.PtrToStructure<KBDLLHOOKSTRUCT>(lParam);
                var vk = (int)data.vkCode;
                bool isDown = wParam == (IntPtr)WM_KEYDOWN;
                bool isUp = wParam == (IntPtr)WM_KEYUP;

                // Track modifiers
                if (vk == 0x11) { _isCtrl = isDown ? true : (isUp ? false : _isCtrl); }
                else if (vk == 0x10) { _isShift = isDown ? true : (isUp ? false : _isShift); }
                else if (vk == 0x12) { _isAlt = isDown ? true : (isUp ? false : _isAlt); }
                else if (vk == 0x5B || vk == 0x5C) { _isWin = isDown ? true : (isUp ? false : _isWin); } // LWIN/RWIN
                else if (isDown || isUp)
                {
                    HandleKeySequenceEvent(vk, isDown);
                }
            }
            return CallNextHookEx(_keyboardHook, nCode, wParam, lParam);
        }

        // --------- KeySequence 錄製狀態 ---------
        private KeySequenceStep? _pendingSeq;
        private DateTime _lastKeyEventTime = DateTime.Now;
        private const int SequenceSplitIdleMs = 600; // 超過此閒置視為新序列

        private void HandleKeySequenceEvent(int vk, bool down)
        {
            string disp = VkToDisplay(vk);
            // 檢查是否是控制熱鍵 (只對 keydown 做過濾)
            if (down)
            {
                var parts = new System.Collections.Generic.List<string>();
                if (_isCtrl) parts.Add("Ctrl");
                if (_isShift) parts.Add("Shift");
                if (_isAlt) parts.Add("Alt");
                if (_isWin) parts.Add("Win");
                parts.Add(disp);
                var comboStr = string.Join("+", parts);
                if (ReservedCombos.Contains(comboStr))
                {
                    Log("[Recorder] (skip control hotkey) " + comboStr);
                    _lastKeyEventTime = DateTime.Now; // 不產生事件但刷新時間避免爆量 Delay
                    return;
                }
            }

            var now = DateTime.Now;
            int gap = (int)(now - _lastKeyEventTime).TotalMilliseconds;
            _lastKeyEventTime = now;

            // 判斷是否要開啟新序列
            if (_pendingSeq == null || gap > SequenceSplitIdleMs)
            {
                // 先將舊序列提交
                FinalizePendingSequence();
                AddDelayFrom(_lastEventTime); // 序列前保持與其他類事件的 Delay 步驟一致
                _pendingSeq = new KeySequenceStep();
            }
            _pendingSeq.Events.Add(new KeyEventItem { Key = disp, Down = down, DelayMsBefore = Math.Max(0, gap) });
        }

        private void FinalizePendingSequence()
        {
            if (_pendingSeq == null) return;
            if (_pendingSeq.Events.Count > 0)
            {
                _steps.Add(_pendingSeq);
                var preview = string.Join(" ", _pendingSeq.Events
                    .Take(6)
                    .Select(e => (e.Down ? "↓" : "↑") + e.Key));
                if (_pendingSeq.Events.Count > 6) preview += " ...";
                var line = $"KeySequence { _pendingSeq.Events.Count } ev: {preview}";
                Log("[Recorder] +" + line);
                StepCaptured?.Invoke(_pendingSeq, line);
            }
            _pendingSeq = null;
        }

    // 移除重複定義 (上方已整合序列 finalize)

        private static string VkToDisplay(int vk)
        {
            // Handle letters, digits, and function keys
            // First map common special keys to friendly names that our InputService understands
            switch (vk)
            {
                case 0x08: return "Backspace"; // VK_BACK
                case 0x09: return "Tab";       // VK_TAB
                case 0x0D: return "Enter";     // VK_RETURN
                case 0x1B: return "Esc";       // VK_ESCAPE
                case 0x20: return "Space";     // VK_SPACE
                case 0x21: return "PageUp";    // VK_PRIOR
                case 0x22: return "PageDown";  // VK_NEXT
                case 0x23: return "End";       // VK_END
                case 0x24: return "Home";      // VK_HOME
                case 0x25: return "Left";      // VK_LEFT
                case 0x26: return "Up";        // VK_UP
                case 0x27: return "Right";     // VK_RIGHT
                case 0x28: return "Down";      // VK_DOWN
                case 0x2D: return "Insert";    // VK_INSERT
                case 0x2E: return "Delete";    // VK_DELETE
            }

            if (vk >= 0x30 && vk <= 0x39) return ((char)vk).ToString();
            if (vk >= 0x41 && vk <= 0x5A) return ((char)vk).ToString();
            if (vk >= 0x70 && vk <= 0x87) return $"F{vk - 0x6F}"; // F1..F24
            // Fall back to VK_XX hex representation (will be parsed by InputService with hex support)
            return $"VK_{vk:X2}";
        }

        private void Log(string msg) => OnLog?.Invoke(msg);

        public void Dispose()
        {
            try
            {
                if (_mouseHook != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(_mouseHook);
                    _mouseHook = IntPtr.Zero;
                }
            }
            catch { }
        }
    }
}
