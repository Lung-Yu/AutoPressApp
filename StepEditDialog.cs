using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using AutoPressApp.Steps;

namespace AutoPressApp
{
    public class StepEditDialog : Form
    {
        private Step _step; // reference edited directly
        private Panel panel;        
        private Button btnOk;        
        private Button btnCancel;
        private Control? primaryFocus;

        public StepEditDialog(Step step)
        {
            _step = step;
            Text = "編輯步驟";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = MinimizeBox = false;
            ClientSize = new Size(360, 260);

            panel = new Panel { Dock = DockStyle.Top, Height = 200 };            
            btnOk = new Button { Text = "確定", DialogResult = DialogResult.OK, Left = 190, Width = 70, Top = 210 };
            btnCancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Left = 270, Width = 70, Top = 210 };
            Controls.Add(panel);
            Controls.Add(btnOk);
            Controls.Add(btnCancel);

            BuildUI();

            AcceptButton = btnOk;
            CancelButton = btnCancel;
            Shown += (_, __) => primaryFocus?.Focus();
            btnOk.Click += (_, __) => { if (!ApplyChanges()) this.DialogResult = DialogResult.None; };
        }

        private void BuildUI()
        {
            panel.Controls.Clear();
            int y = 15;
            void Label(string text)
            {
                panel.Controls.Add(new Label { Text = text, Left = 15, Top = y + 4, AutoSize = true });
            }
            Control Input(Control c)
            {
                c.Left = 120; c.Top = y; c.Width = 200; panel.Controls.Add(c); y += 32; return c;
            }

            switch (_step)
            {
                case DelayStep d:
                    Label("延遲 (ms)"); primaryFocus = Input(new NumericUpDown { Minimum = 0, Maximum = 2147483647, Value = d.Ms, Name = "numDelay" });
                    break;
                case MouseClickStep m:
                    Label("X"); var numX = (NumericUpDown)Input(new NumericUpDown { Minimum = -10000, Maximum = 10000, Value = m.X, Name = "numX" });
                    Label("Y"); var numY = (NumericUpDown)Input(new NumericUpDown { Minimum = -10000, Maximum = 10000, Value = m.Y, Name = "numY" });
                    Label("按鍵"); var cmbBtn = (ComboBox)Input(new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Name = "cmbBtn", Items = { "left", "right", "middle" } });
                    cmbBtn.SelectedItem = m.Button;
                    break;
                case FocusWindowStep f:
                    Label("視窗標題包含"); primaryFocus = Input(new TextBox { Text = f.TitleContains ?? string.Empty, Name = "txtTitle" });
                    Label("行程名稱(Process)"); Input(new TextBox { Text = f.ProcessName ?? string.Empty, Name = "txtProc" });
                    Label("等待逾時(ms)"); Input(new NumericUpDown { Minimum = 0, Maximum = 60000, Value = f.TimeoutMs, Name = "numTimeout" });
                    break;
                case LogStep lg:
                    Label("訊息"); primaryFocus = Input(new TextBox { Text = lg.Message, Name = "txtMsg" });
                    break;
                case KeyComboStep kc:
                    Label("組合鍵 (以 + 分隔)"); primaryFocus = Input(new TextBox { Text = string.Join('+', kc.Keys), Name = "txtKeys" });
                    break;
                case KeySequenceStep ks:
                    // Read-only viewer
                    var box = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, ReadOnly = true, Height = 170, Width = 320, Left = 15, Top = 15 };
                    box.Text = string.Join(Environment.NewLine, ks.Events.Select((e,i)=>$"{i:D3} {(e.Down?"Down":"Up  ")} {e.Key} (+{e.DelayMsBefore}ms)"));
                    panel.Controls.Add(box);
                    primaryFocus = box;
                    y = 190;
                    Controls.Add(new Label { Text = "(目前不支援直接編輯按鍵序列)", Left = 15, Top = 190, AutoSize = true, ForeColor = Color.DarkGoldenrod });
                    break;
                default:
                    panel.Controls.Add(new Label { Text = "此步驟類型暫不支援編輯", Left = 15, Top = 20, AutoSize = true, ForeColor = Color.DarkRed });
                    break;
            }
        }

        private bool ApplyChanges()
        {
            try
            {
                switch (_step)
                {
                    case DelayStep d:
                        var numDelay = panel.Controls.Find("numDelay", true).FirstOrDefault() as NumericUpDown; if (numDelay != null) d.Ms = (int)numDelay.Value; break;
                    case MouseClickStep m:
                        var numX = panel.Controls.Find("numX", true).FirstOrDefault() as NumericUpDown; if (numX != null) m.X = (int)numX.Value;
                        var numY = panel.Controls.Find("numY", true).FirstOrDefault() as NumericUpDown; if (numY != null) m.Y = (int)numY.Value;
                        var cmbBtn = panel.Controls.Find("cmbBtn", true).FirstOrDefault() as ComboBox; if (cmbBtn != null && cmbBtn.SelectedItem is string s) m.Button = s;
                        break;
                    case FocusWindowStep f:
                        var txtTitle = panel.Controls.Find("txtTitle", true).FirstOrDefault() as TextBox; if (txtTitle != null) f.TitleContains = string.IsNullOrWhiteSpace(txtTitle.Text) ? null : txtTitle.Text;
                        var txtProc = panel.Controls.Find("txtProc", true).FirstOrDefault() as TextBox; if (txtProc != null) f.ProcessName = string.IsNullOrWhiteSpace(txtProc.Text) ? null : txtProc.Text;
                        var numTimeout = panel.Controls.Find("numTimeout", true).FirstOrDefault() as NumericUpDown; if (numTimeout != null) f.TimeoutMs = (int)numTimeout.Value;
                        break;
                    case LogStep lg:
                        var txtMsg = panel.Controls.Find("txtMsg", true).FirstOrDefault() as TextBox; if (txtMsg != null) lg.Message = txtMsg.Text;
                        break;
                    case KeyComboStep kc:
                        var txtKeys = panel.Controls.Find("txtKeys", true).FirstOrDefault() as TextBox; if (txtKeys != null)
                        {
                            kc.Keys = txtKeys.Text.Split('+', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                        }
                        break;
                    case KeySequenceStep:
                        // read-only
                        break;
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("套用變更失敗: " + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
    }
}
