namespace SoftwareInstaller.UI
{
    partial class SoftwareEditForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.nameLabel = new System.Windows.Forms.Label();
            this.nameTextBox = new System.Windows.Forms.TextBox();
            this.versionLabel = new System.Windows.Forms.Label();
            this.versionTextBox = new System.Windows.Forms.TextBox();
            this.descriptionLabel = new System.Windows.Forms.Label();
            this.descriptionTextBox = new System.Windows.Forms.TextBox();
            this.filePathLabel = new System.Windows.Forms.Label();
            this.filePathTextBox = new System.Windows.Forms.TextBox();
            this.browseButton = new System.Windows.Forms.Button();
            this.silentArgsLabel = new System.Windows.Forms.Label();
            this.silentArgsTextBox = new System.Windows.Forms.TextBox();
            
            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // nameLabel
            // 
            this.nameLabel.AutoSize = true;
            this.nameLabel.Location = new System.Drawing.Point(12, 15);
            this.nameLabel.Name = "nameLabel";
            this.nameLabel.Size = new System.Drawing.Size(59, 15);
            this.nameLabel.Text = "软件名称";
            // 
            // nameTextBox
            // 
            this.nameTextBox.Location = new System.Drawing.Point(90, 12);
            this.nameTextBox.Size = new System.Drawing.Size(220, 23);
            // 
            // versionLabel
            // 
            this.versionLabel.AutoSize = true;
            this.versionLabel.Location = new System.Drawing.Point(12, 45);
            this.versionLabel.Name = "versionLabel";
            this.versionLabel.Size = new System.Drawing.Size(32, 15);
            this.versionLabel.Text = "版本";
            // 
            // versionTextBox
            // 
            this.versionTextBox.Location = new System.Drawing.Point(90, 42);
            this.versionTextBox.Size = new System.Drawing.Size(220, 23);
            // 
            // descriptionLabel
            // 
            this.descriptionLabel.AutoSize = true;
            this.descriptionLabel.Location = new System.Drawing.Point(12, 75);
            this.descriptionLabel.Name = "descriptionLabel";
            this.descriptionLabel.Size = new System.Drawing.Size(32, 15);
            this.descriptionLabel.Text = "说明";
            // 
            // descriptionTextBox
            // 
            this.descriptionTextBox.Location = new System.Drawing.Point(90, 72);
            this.descriptionTextBox.Size = new System.Drawing.Size(220, 23);
            // 
            // filePathLabel
            // 
            this.filePathLabel.AutoSize = true;
            this.filePathLabel.Location = new System.Drawing.Point(12, 105);
            this.filePathLabel.Name = "filePathLabel";
            this.filePathLabel.Size = new System.Drawing.Size(56, 15);
            this.filePathLabel.Text = "文件路径";
            // 
            // filePathTextBox
            // 
            this.filePathTextBox.Location = new System.Drawing.Point(90, 102);
            this.filePathTextBox.Size = new System.Drawing.Size(150, 23);
            // 
            // browseButton
            // 
            this.browseButton.Location = new System.Drawing.Point(250, 102);
            this.browseButton.Size = new System.Drawing.Size(60, 23);
            this.browseButton.Text = "浏览...";
            this.browseButton.Click += new System.EventHandler(this.browseButton_Click);
            // 
            // silentArgsLabel
            // 
            this.silentArgsLabel.AutoSize = true;
            this.silentArgsLabel.Location = new System.Drawing.Point(12, 135);
            this.silentArgsLabel.Name = "silentArgsLabel";
            this.silentArgsLabel.Size = new System.Drawing.Size(80, 15);
            this.silentArgsLabel.Text = "静默安装参数";
            // 
            // silentArgsTextBox
            // 
            this.silentArgsTextBox.Location = new System.Drawing.Point(90, 132);
            this.silentArgsTextBox.Size = new System.Drawing.Size(220, 23);
            
            
            // 
            // okButton
            // 
            this.okButton.Location = new System.Drawing.Point(120, 170);
            this.okButton.Size = new System.Drawing.Size(75, 28);
            this.okButton.Text = "确定";
            this.okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            // 
            // cancelButton
            // 
            this.cancelButton.Location = new System.Drawing.Point(210, 170);
            this.cancelButton.Size = new System.Drawing.Size(75, 28);
            this.cancelButton.Text = "取消";
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            // 
            // SoftwareEditForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(330, 210);
            this.Controls.Add(this.nameLabel);
            this.Controls.Add(this.nameTextBox);
            this.Controls.Add(this.versionLabel);
            this.Controls.Add(this.versionTextBox);
            this.Controls.Add(this.descriptionLabel);
            this.Controls.Add(this.descriptionTextBox);
            this.Controls.Add(this.filePathLabel);
            this.Controls.Add(this.filePathTextBox);
            this.Controls.Add(this.browseButton);
            this.Controls.Add(this.silentArgsLabel);
            this.Controls.Add(this.silentArgsTextBox);
            
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.cancelButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SoftwareEditForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "编辑软件";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.Label nameLabel;
        private System.Windows.Forms.TextBox nameTextBox;
        private System.Windows.Forms.Label versionLabel;
        private System.Windows.Forms.TextBox versionTextBox;
        private System.Windows.Forms.Label descriptionLabel;
        private System.Windows.Forms.TextBox descriptionTextBox;
        private System.Windows.Forms.Label filePathLabel;
        private System.Windows.Forms.TextBox filePathTextBox;
        private System.Windows.Forms.Button browseButton;
        private System.Windows.Forms.Label silentArgsLabel;
        private System.Windows.Forms.TextBox silentArgsTextBox;
        
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Button cancelButton;
    }
}