using System;
using System.Drawing;
using System.Windows.Forms;
using AutoPressApp.Steps;

namespace AutoPressApp
{
    public class KeyEventEditDialog : Form
    {
        private KeyEventItem _item;
        private TextBox txtKey;
        private CheckBox chkDown;
        private NumericUpDown numDelay;
        private Button btnOk;
        private Button btnCancel;
        public KeyEventEditDialog(KeyEventItem item, int index)
        {
            _item = item;
            Text = $"編輯事件 #{index}";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = MinimizeBox = false;
            ClientSize = new Size(300, 170);

            Controls.Add(new Label { Text = "按鍵(Key)", Left = 15, Top = 18, AutoSize = true });
            txtKey = new TextBox { Left = 110, Top = 15, Width = 160, Text = item.Key };
            Controls.Add(txtKey);

            Controls.Add(new Label { Text = "按下?", Left = 15, Top = 56, AutoSize = true });
            chkDown = new CheckBox { Left = 110, Top = 54, Checked = item.Down };
            Controls.Add(chkDown);

            Controls.Add(new Label { Text = "延遲(ms)", Left = 15, Top = 92, AutoSize = true });
            numDelay = new NumericUpDown { Left = 110, Top = 90, Width = 100, Minimum = 0, Maximum = 1000000, Value = item.DelayMsBefore };
            Controls.Add(numDelay);

            btnOk = new Button { Text = "確定", DialogResult = DialogResult.OK, Left = 110, Top = 125, Width = 70 };
            btnCancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Left = 200, Top = 125, Width = 70 };
            Controls.AddRange(new Control[] { btnOk, btnCancel });

            AcceptButton = btnOk; CancelButton = btnCancel;
            btnOk.Click += (s, e) => { if (!Apply()) DialogResult = DialogResult.None; };
        }

        private bool Apply()
        {
            string key = txtKey.Text.Trim();
            if (string.IsNullOrEmpty(key))
            {
                MessageBox.Show("Key 不可為空", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            _item.Key = key;
            _item.Down = chkDown.Checked;
            _item.DelayMsBefore = (int)numDelay.Value;
            return true;
        }
    }
}
