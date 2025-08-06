using SoftwareInstaller.Core;
using SoftwareInstaller.Models;
using SoftwareInstaller.Utils;
using System;
using System.Windows.Forms;

namespace SoftwareInstaller.UI
{
    public partial class SoftwareEditForm : Form
    {
        public SoftwareItem SoftwareItem { get; private set; }

        public SoftwareEditForm(SoftwareItem? item = null)
        {
            InitializeComponent();
            SoftwareItem = item ?? new SoftwareItem
            {
                Name = string.Empty,
                Version = string.Empty,
                Size = string.Empty,
                Description = string.Empty,
                Category = string.Empty
            };
            if (item != null)
            {
                this.Text = "编辑软件";
                nameTextBox.Text = item.Name;
                versionTextBox.Text = item.Version;
                descriptionTextBox.Text = item.Description;
                filePathTextBox.Text = item.FilePath;
                silentArgsTextBox.Text = item.SilentInstallArgs;
            }
            else
            {
                this.Text = "新增软件";
            }
        }

        private void browseButton_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "可执行文件 (*.exe)|*.exe|所有文件 (*.*)|*.*";
                openFileDialog.Title = "请选择软件安装包";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePathTextBox.Text = openFileDialog.FileName;
                    string? productName = FileMetadataReader.GetProductName(openFileDialog.FileName);
                    if (!string.IsNullOrWhiteSpace(productName))
                    {
                        nameTextBox.Text = productName;
                    }
                }
            }
        }

        
    }
}
