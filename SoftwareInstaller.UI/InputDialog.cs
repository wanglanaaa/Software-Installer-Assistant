
using System.Drawing;
using System.Windows.Forms;

namespace SoftwareInstaller.UI
{
    public class InputDialog : Form
    {
        private Label promptLabel;
        private TextBox inputTextBox;
        private Button okButton;
        private Button cancelButton;

        public string InputText { get; private set; } = string.Empty;

        public InputDialog(string title, string prompt)
        {
            this.Text = title;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.ClientSize = new Size(380, 120);
            this.ControlBox = false;

            promptLabel = new Label()
            {
                Text = prompt,
                Location = new Point(20, 20),
                AutoSize = true,
                Font = new Font("Segoe UI", 9F)
            };

            inputTextBox = new TextBox()
            {
                Location = new Point(20, 50),
                Size = new Size(340, 23),
                Font = new Font("Segoe UI", 9F)
            };

            okButton = new Button()
            {
                Text = "确定",
                DialogResult = DialogResult.OK,
                Location = new Point(200, 80),
                Size = new Size(75, 25)
            };

            cancelButton = new Button()
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Location = new Point(285, 80),
                Size = new Size(75, 25)
            };

            this.Controls.Add(promptLabel);
            this.Controls.Add(inputTextBox);
            this.Controls.Add(okButton);
            this.Controls.Add(cancelButton);

            this.AcceptButton = okButton;
            this.CancelButton = cancelButton;

            okButton.Click += (sender, e) => {
                InputText = inputTextBox.Text;
                this.DialogResult = DialogResult.OK;
            };
        }
    }
}
