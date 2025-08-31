using System;
using System.Drawing;
using System.Windows.Forms;

namespace AutoPressApp
{
    public class InputBoxDialog : Form
    {
        private readonly TextBox _text;
        private readonly Label _label;
        private readonly Button _ok;
        private readonly Button _cancel;
        public string ResultText => _text.Text.Trim();

        public InputBoxDialog(string title, string prompt, string? defaultValue = null)
        {
            Text = title;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(380, 140);
            MaximizeBox = MinimizeBox = false;
            ShowInTaskbar = false;

            _label = new Label { Left = 15, Top = 15, Width = 350, Text = prompt };
            _text = new TextBox { Left = 15, Top = 45, Width = 350 };
            if (!string.IsNullOrWhiteSpace(defaultValue)) _text.Text = defaultValue;
            _ok = new Button { Text = "確定", DialogResult = DialogResult.OK, Left = 190, Width = 80, Top = 85 };
            _cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Left = 285, Width = 80, Top = 85 };
            AcceptButton = _ok;
            CancelButton = _cancel;

            Controls.AddRange(new Control[] { _label, _text, _ok, _cancel });
        }
    }
}
