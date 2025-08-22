using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace AutoPressApp.Services
{
    public class WindowService
    {
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")] private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hWnd);
        private const int SW_RESTORE = 9;

        public bool FocusWindow(string? titleContains, string? processName, int timeoutMs)
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                foreach (var p in Process.GetProcesses().Where(p => p.MainWindowHandle != IntPtr.Zero && IsWindowVisible(p.MainWindowHandle)))
                {
                    if (!string.IsNullOrEmpty(processName) && !string.Equals(p.ProcessName, processName, StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (!string.IsNullOrEmpty(titleContains) && (p.MainWindowTitle?.IndexOf(titleContains, StringComparison.OrdinalIgnoreCase) ?? -1) < 0)
                        continue;

                    var h = p.MainWindowHandle;
                    ShowWindow(h, SW_RESTORE);
                    SetForegroundWindow(h);
                    return true;
                }
                System.Threading.Thread.Sleep(200);
            }
            return false;
        }
    }
}
