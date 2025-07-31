using System;
using System.Drawing;
using System.Windows.Forms;

namespace SoftwareInstaller.UI
{
    public class InputDialog : Form
    {
        private TextBox inputTextBox;
        private Button okButton;
        private Button cancelButton;
        public string? InputText => inputTextBox.Text;

        public InputDialog(string title, string prompt)
        {
            this.Text = title;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ClientSize = new Size(350, 140);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;
            this.Font = new Font("Segoe UI", 9F);

            Label promptLabel = new Label
            {
                Text = prompt,
                AutoSize = false,
                Location = new Point(15, 15),
                Size = new Size(320, 30)
            };
            inputTextBox = new TextBox
            {
                Location = new Point(15, 50),
                Size = new Size(320, 23)
            };
            okButton = new Button
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Location = new Point(170, 90),
                Size = new Size(75, 28)
            };
            cancelButton = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Location = new Point(260, 90),
                Size = new Size(75, 28)
            };

            this.Controls.Add(promptLabel);
            this.Controls.Add(inputTextBox);
            this.Controls.Add(okButton);
            this.Controls.Add(cancelButton);

            this.AcceptButton = okButton;
            this.CancelButton = cancelButton;
        }
    }
}