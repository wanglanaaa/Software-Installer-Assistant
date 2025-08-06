namespace SoftwareInstaller.UI;

partial class MainForm
{
    /// <summary>
    ///  Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    ///  Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    ///  设计器支持所需的方法 - 不要修改
    ///  使用代码编辑器修改此方法的内容。
    /// </summary>
    private void InitializeComponent()
    {
            this.categoryTreeView = new System.Windows.Forms.TreeView();
            this.softwareListView = new System.Windows.Forms.ListView();
            this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
            this.columnHeader2 = new System.Windows.Forms.ColumnHeader();
            this.columnHeader3 = new System.Windows.Forms.ColumnHeader();
            this.columnHeader4 = new System.Windows.Forms.ColumnHeader();
            this.columnHeader5 = new System.Windows.Forms.ColumnHeader();
            this.addCategoryButton = new System.Windows.Forms.Button();
            this.deleteCategoryButton = new System.Windows.Forms.Button();
            this.addSoftwareButton = new System.Windows.Forms.Button();
            this.deleteSoftwareButton = new System.Windows.Forms.Button();
            this.bottomPanel = new System.Windows.Forms.Panel();
            this.installPathTextBox = new System.Windows.Forms.TextBox();
            this.browseButton = new System.Windows.Forms.Button();
            this.manualInstallButton = new System.Windows.Forms.Button();
            this.autoInstallButton = new System.Windows.Forms.Button();
            this.schemeComboBox = new System.Windows.Forms.ComboBox();
            this.saveSchemeButton = new System.Windows.Forms.Button();
            this.deleteSchemeButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // categoryTreeView
            // 
            this.categoryTreeView.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.categoryTreeView.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.categoryTreeView.CheckBoxes = true;
            this.categoryTreeView.Location = new System.Drawing.Point(20, 20);
            this.categoryTreeView.Name = "categoryTreeView";
            this.categoryTreeView.ShowLines = false;
            this.categoryTreeView.Size = new System.Drawing.Size(200, 426);
            this.categoryTreeView.TabIndex = 0;
            // 
            // softwareListView
            // 
            this.softwareListView.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.softwareListView.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader3,
            this.columnHeader4,
            this.columnHeader5});
            this.softwareListView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.softwareListView.FullRowSelect = true;
            this.softwareListView.Location = new System.Drawing.Point(230, 20);
            this.softwareListView.Name = "softwareListView";
            this.softwareListView.Size = new System.Drawing.Size(542, 490);
            this.softwareListView.TabIndex = 1;
            this.softwareListView.UseCompatibleStateImageBehavior = false;
            this.softwareListView.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "序号";
            this.columnHeader1.Width = 50;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "软件名称";
            this.columnHeader2.Width = 150;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "版本";
            this.columnHeader3.Width = 80;
            // 
            // columnHeader4
            // 
            this.columnHeader4.Text = "大小";
            this.columnHeader4.Width = 80;
            // 
            // columnHeader5
            // 
            this.columnHeader5.Text = "说明";
            this.columnHeader5.Width = 180;
            // 
            // addCategoryButton
            // 
            this.addCategoryButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.addCategoryButton.AutoSize = true;
            this.addCategoryButton.Location = new System.Drawing.Point(16, 456);
            this.addCategoryButton.Name = "addCategoryButton";
            this.addCategoryButton.Size = new System.Drawing.Size(99, 25);
            this.addCategoryButton.TabIndex = 2;
            this.addCategoryButton.Text = "➕ 新增分类";
            this.addCategoryButton.UseVisualStyleBackColor = true;
            // 
            // deleteCategoryButton
            // 
            this.deleteCategoryButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.deleteCategoryButton.AutoSize = true;
            this.deleteCategoryButton.Location = new System.Drawing.Point(121, 456);
            this.deleteCategoryButton.Name = "deleteCategoryButton";
            this.deleteCategoryButton.Size = new System.Drawing.Size(99, 25);
            this.deleteCategoryButton.TabIndex = 3;
            this.deleteCategoryButton.Text = "➖ 删除分类";
            this.deleteCategoryButton.UseVisualStyleBackColor = true;
            // 
            // addSoftwareButton
            // 
            this.addSoftwareButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.addSoftwareButton.AutoSize = true;
            this.addSoftwareButton.Location = new System.Drawing.Point(16, 486);
            this.addSoftwareButton.Name = "addSoftwareButton";
            this.addSoftwareButton.Size = new System.Drawing.Size(99, 25);
            this.addSoftwareButton.TabIndex = 4;
            this.addSoftwareButton.Text = "➕ 新增软件";
            this.addSoftwareButton.UseVisualStyleBackColor = true;
            // 
            // deleteSoftwareButton
            // 
            this.deleteSoftwareButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.deleteSoftwareButton.AutoSize = true;
            this.deleteSoftwareButton.Location = new System.Drawing.Point(121, 486);
            this.deleteSoftwareButton.Name = "deleteSoftwareButton";
            this.deleteSoftwareButton.Size = new System.Drawing.Size(99, 25);
            this.deleteSoftwareButton.TabIndex = 5;
            this.deleteSoftwareButton.Text = "➖ 删除软件";
            this.deleteSoftwareButton.UseVisualStyleBackColor = true;
            // 
            // bottomPanel
            // 
            this.bottomPanel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.bottomPanel.BackColor = System.Drawing.Color.White;
            this.bottomPanel.Location = new System.Drawing.Point(0, 522);
            this.bottomPanel.Name = "bottomPanel";
            this.bottomPanel.Size = new System.Drawing.Size(784, 80);
            this.bottomPanel.TabIndex = 6;
            //
            // Note: The controls inside bottomPanel are added dynamically in MainForm.cs
            // for horizontal centering, so they are not added to bottomPanel.Controls here.
            // They are just declared and instantiated.
            // 
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(784, 601);
            this.Controls.Add(this.bottomPanel);
            this.Controls.Add(this.deleteSoftwareButton);
            this.Controls.Add(this.addSoftwareButton);
            this.Controls.Add(this.deleteCategoryButton);
            this.Controls.Add(this.addCategoryButton);
            this.Controls.Add(this.softwareListView);
            this.Controls.Add(this.categoryTreeView);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MinimumSize = new System.Drawing.Size(800, 640);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "软件安装助理";
            this.ResumeLayout(false);
            this.PerformLayout();

    }

    #endregion

    private TreeView categoryTreeView;
    private ListView softwareListView;
    private Button addCategoryButton;
    private Button deleteCategoryButton;
    private Button addSoftwareButton;
    private Button deleteSoftwareButton;
    private Panel bottomPanel;
    private TextBox installPathTextBox;
    private Button browseButton;
    private Button manualInstallButton;
    private Button autoInstallButton;
    private ComboBox schemeComboBox;
    private Button saveSchemeButton;
    private Button deleteSchemeButton;
    private ColumnHeader columnHeader1;
    private ColumnHeader columnHeader2;
    private ColumnHeader columnHeader3;
    private ColumnHeader columnHeader4;
private System.Windows.Forms.ColumnHeader columnHeader5;
}
