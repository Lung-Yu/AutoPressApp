using System;
using System.Globalization;
using System.Runtime.InteropServices;
using AutoPressApp.Steps;

namespace AutoPressApp.Services
{
    public class InputService
    {
        [DllImport("user32.dll")] static extern bool SetCursorPos(int X, int Y);
        [DllImport("user32.dll")] static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
    [DllImport("user32.dll")] static extern short VkKeyScan(char ch);
    [DllImport("user32.dll")] static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo); // legacy fallback
    [DllImport("user32.dll", SetLastError = true)] static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        private const uint MOUSEEVENTF_RIGHTUP = 0x0010;
    private const int KEYEVENTF_KEYUP = 0x0002;
    private const uint INPUT_KEYBOARD = 1;
    private const uint KEYEVENTF_SCANCODE = 0x0008;
    private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;

    [StructLayout(LayoutKind.Sequential)]
    struct INPUT { public uint type; public InputUnion U; }
    [StructLayout(LayoutKind.Explicit)]
    struct InputUnion { [FieldOffset(0)] public MOUSEINPUT mi; [FieldOffset(0)] public KEYBDINPUT ki; [FieldOffset(0)] public HARDWAREINPUT hi; }
    [StructLayout(LayoutKind.Sequential)] struct MOUSEINPUT { public int dx; public int dy; public uint mouseData; public uint dwFlags; public uint time; public UIntPtr dwExtraInfo; }
    [StructLayout(LayoutKind.Sequential)] struct KEYBDINPUT { public ushort wVk; public ushort wScan; public uint dwFlags; public uint time; public UIntPtr dwExtraInfo; }
    [StructLayout(LayoutKind.Sequential)] struct HARDWAREINPUT { public uint uMsg; public ushort wParamL; public ushort wParamH; }
    [DllImport("user32.dll")] static extern uint MapVirtualKey(uint uCode, uint uMapType);

        public void Click(int x, int y, string button, CoordMode mode)
        {
            // For now, only screen coords supported; relative-to-window can be added later
            SetCursorPos(x, y);
            if (string.Equals(button, "right", StringComparison.OrdinalIgnoreCase))
            {
                mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, UIntPtr.Zero);
            }
            else
            {
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
            }
        }

    public void SendKeyCombo(string[] keys) => SendKeyCombo(keys, holdMs: 0, preDelayMs: 0);

    public void SendKeyCombo(string[] keys, int holdMs, int preDelayMs)
        {
            // Parse modifiers and primary key (last non-modifier)
            bool ctrl = false, shift = false, alt = false, win = false;
            string? primary = null;
            foreach (var raw in keys)
            {
                var s = raw.Trim();
                if (s.Equals("Ctrl", StringComparison.OrdinalIgnoreCase)) ctrl = true;
                else if (s.Equals("Shift", StringComparison.OrdinalIgnoreCase)) shift = true;
                else if (s.Equals("Alt", StringComparison.OrdinalIgnoreCase)) alt = true;
                else if (s.Equals("Win", StringComparison.OrdinalIgnoreCase) || s.Equals("Windows", StringComparison.OrdinalIgnoreCase)) win = true;
                else primary = s; // last wins
            }

            // Build ordered list: modifiers down, primary down/up, modifiers up (reverse order)
            if (preDelayMs > 0) System.Threading.Thread.Sleep(preDelayMs);

            var seq = new System.Collections.Generic.List<INPUT>();
            void addVkKey(ushort vk, bool down)
            {
                if (vk == 0) return;
                uint flags = down ? 0u : KEYEVENTF_KEYUP;
                // Extended keys (arrows, insert/delete/home/end/page) need EXTENDED flag when using VK mode
                if (vk is 0x21 or 0x22 or 0x23 or 0x24 or 0x25 or 0x26 or 0x27 or 0x28 or 0x2D or 0x2E)
                    flags |= KEYEVENTF_EXTENDEDKEY;
                var inp = new INPUT { type = INPUT_KEYBOARD, U = new InputUnion { ki = new KEYBDINPUT { wVk = vk, wScan = (ushort)MapVirtualKey(vk, 0), dwFlags = flags, time = 0, dwExtraInfo = UIntPtr.Zero } } };
                seq.Add(inp);
            }
            void addScanKey(ushort vk, bool down)
            {
                if (vk == 0) return;
                var scan = (ushort)MapVirtualKey(vk, 0); // MAPVK_VK_TO_VSC
                if (scan == 0) { addVkKey(vk, down); return; }
                uint flags = KEYEVENTF_SCANCODE | (down ? 0u : KEYEVENTF_KEYUP);
                // Extended keys require EXTENDED flag even in scancode mode for arrows etc.
                if (vk is 0x21 or 0x22 or 0x23 or 0x24 or 0x25 or 0x26 or 0x27 or 0x28 or 0x2D or 0x2E)
                    flags |= KEYEVENTF_EXTENDEDKEY;
                var inp = new INPUT { type = INPUT_KEYBOARD, U = new InputUnion { ki = new KEYBDINPUT { wVk = 0, wScan = scan, dwFlags = flags, time = 0, dwExtraInfo = UIntPtr.Zero } } };
                seq.Add(inp);
            }

            if (ctrl) addVkKey(0x11, true);
            if (shift) addVkKey(0x10, true);
            if (alt) addVkKey(0x12, true);
            if (win) addVkKey(0x5B, true);

            ushort primaryVk = 0;
            if (!string.IsNullOrEmpty(primary))
            {
                primaryVk = MapKeyStringToVk(primary!);
                if (primaryVk != 0)
                {
                    addScanKey(primaryVk, true); // key down
                }
            }

            // Send modifiers down + primary down first
            if (seq.Count > 0)
            {
                try
                {
                    var sent = SendInput((uint)seq.Count, seq.ToArray(), Marshal.SizeOf<INPUT>());
                    if (sent != (uint)seq.Count) FallbackSend(ctrl, shift, alt, win, null);
                }
                catch { FallbackSend(ctrl, shift, alt, win, null); }
            }

            if (primaryVk != 0)
            {
                if (holdMs > 0) System.Threading.Thread.Sleep(holdMs);
                // release primary
                var relList = new System.Collections.Generic.List<INPUT>();
                addScanKey(primaryVk, false);
                // we appended release to seq; but we don't want duplicates so build a dedicated release list for primary
                relList.Add(seq[^1]); // last added is primary up
                seq.RemoveAt(seq.Count - 1);
                try { SendInput((uint)relList.Count, relList.ToArray(), Marshal.SizeOf<INPUT>()); } catch { /* ignore */ }
            }

            // Release modifiers in reverse order (separate to ensure holdMs doesn't hold them unintentionally)
            var modsRelease = new System.Collections.Generic.List<INPUT>();
            void addModRelease(ushort vk)
            {
                if (vk == 0) return; var scan = (ushort)MapVirtualKey(vk, 0); uint flags = KEYEVENTF_SCANCODE | KEYEVENTF_KEYUP; if (vk is 0x21 or 0x22 or 0x23 or 0x24 or 0x25 or 0x26 or 0x27 or 0x28 or 0x2D or 0x2E) flags |= KEYEVENTF_EXTENDEDKEY; if (scan == 0) { // fallback VK
                    modsRelease.Add(new INPUT { type = INPUT_KEYBOARD, U = new InputUnion { ki = new KEYBDINPUT { wVk = vk, wScan = (ushort)MapVirtualKey(vk,0), dwFlags = KEYEVENTF_KEYUP | ((vk is 0x21 or 0x22 or 0x23 or 0x24 or 0x25 or 0x26 or 0x27 or 0x28 or 0x2D or 0x2E)?KEYEVENTF_EXTENDEDKEY:0), time = 0, dwExtraInfo = UIntPtr.Zero } } });
                }
                else
                {
                    modsRelease.Add(new INPUT { type = INPUT_KEYBOARD, U = new InputUnion { ki = new KEYBDINPUT { wVk = 0, wScan = scan, dwFlags = flags, time = 0, dwExtraInfo = UIntPtr.Zero } } });
                }
            }
            if (win) addModRelease(0x5B);
            if (alt) addModRelease(0x12);
            if (shift) addModRelease(0x10);
            if (ctrl) addModRelease(0x11);
            if (modsRelease.Count > 0)
            {
                try { SendInput((uint)modsRelease.Count, modsRelease.ToArray(), Marshal.SizeOf<INPUT>()); } catch { /* ignore */ }
            }
        }

        private void FallbackSend(bool ctrl, bool shift, bool alt, bool win, string? primary)
        {
            void down(byte vk) => keybd_event(vk, 0, 0, UIntPtr.Zero);
            void up(byte vk) => keybd_event(vk, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            if (ctrl) down(0x11);
            if (shift) down(0x10);
            if (alt) down(0x12);
            if (win) down(0x5B);
            if (!string.IsNullOrEmpty(primary))
            {
                var vk = MapKeyStringToVk(primary!);
                if (vk != 0) { down(vk); up(vk); }
            }
            if (win) up(0x5B);
            if (alt) up(0x12);
            if (shift) up(0x10);
            if (ctrl) up(0x11);
        }

    public static byte MapKeyStringToVk(string key)
        {
            // Normalize
            key = key.Trim();

            // Common names
            switch (key.ToLowerInvariant())
            {
                case "enter": return 0x0D; // VK_RETURN
                case "esc":
                case "escape": return 0x1B; // VK_ESCAPE
                case "tab": return 0x09; // VK_TAB
                case "space": return 0x20; // VK_SPACE
                case "backspace": return 0x08; // VK_BACK
                case "delete": return 0x2E; // VK_DELETE
                case "insert": return 0x2D; // VK_INSERT
                case "home": return 0x24; // VK_HOME
                case "end": return 0x23; // VK_END
                case "pageup":
                case "prior": return 0x21; // VK_PRIOR
                case "pagedown":
                case "next": return 0x22; // VK_NEXT
                case "up": return 0x26; // VK_UP
                case "down": return 0x28; // VK_DOWN
                case "left": return 0x25; // VK_LEFT
                case "right": return 0x27; // VK_RIGHT
            }

            // Numpad e.g. Num0..Num9
            if (key.StartsWith("Num", StringComparison.OrdinalIgnoreCase) && key.Length == 4 && char.IsDigit(key[3]))
            {
                int d = key[3] - '0';
                return (byte)(0x60 + d); // VK_NUMPAD0 = 0x60
            }

            // Handle function keys and single letters/digits
            if (key.Length == 1)
            {
                char c = key[0];
                if (char.IsLetter(c)) return (byte)char.ToUpperInvariant(c);
                if (char.IsDigit(c)) return (byte)(0x30 + (c - '0'));
            }
            if (key.StartsWith("F", StringComparison.OrdinalIgnoreCase))
            {
                if (int.TryParse(key.Substring(1), out int fn) && fn >= 1 && fn <= 24)
                    return (byte)(0x70 + (fn - 1)); // VK_F1=0x70
            }
            // Allow raw hex form like VK_2E
            if (key.StartsWith("VK_", StringComparison.OrdinalIgnoreCase))
            {
                var hex = key.Substring(3); // after 'VK_'
                if (byte.TryParse(hex, NumberStyles.HexNumber, null, out var hexVk))
                    return hexVk;
            }
            return 0; // fallback no-op
        }

        // 單一鍵 Down/Up 事件 (供 KeySequenceStep 使用)
        public void SendKeyEvent(string key, bool down)
        {
            var vk = MapKeyStringToVk(key);
            if (vk == 0) return;
            try
            {
                var list = new System.Collections.Generic.List<INPUT>();
                ushort uVk = vk;
                var scan = (ushort)MapVirtualKey(uVk, 0);
                uint flags = 0;
                bool isExtended = uVk is 0x21 or 0x22 or 0x23 or 0x24 or 0x25 or 0x26 or 0x27 or 0x28 or 0x2D or 0x2E;
                if (scan != 0)
                {
                    flags |= KEYEVENTF_SCANCODE;
                    if (isExtended) flags |= KEYEVENTF_EXTENDEDKEY;
                }
                if (!down) flags |= KEYEVENTF_KEYUP;
                list.Add(new INPUT
                {
                    type = INPUT_KEYBOARD,
                    U = new InputUnion
                    {
                        ki = new KEYBDINPUT
                        {
                            wVk = scan == 0 ? uVk : (ushort)0,
                            wScan = scan,
                            dwFlags = flags,
                            time = 0,
                            dwExtraInfo = UIntPtr.Zero
                        }
                    }
                });
                SendInput((uint)list.Count, list.ToArray(), Marshal.SizeOf<INPUT>());
            }
            catch
            {
                // Fallback
                keybd_event(vk, 0, down ? 0u : KEYEVENTF_KEYUP, UIntPtr.Zero);
            }
        }
    }
}
