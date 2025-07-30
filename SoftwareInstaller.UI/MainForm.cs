using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;
using SoftwareInstaller.Core;
using SoftwareInstaller.Models;
using SoftwareInstaller.Utils;

namespace SoftwareInstaller.UI
{
    public partial class MainForm : Form
    {
        private TreeView categoryTreeView = null!;
        private ListView softwareListView = null!;
        private TextBox installPathTextBox = null!;
        private Button browseButton = null!;
        private Button manualInstallButton = null!;
        private Button autoInstallButton = null!;
        private Button addCategoryButton = null!;
        private Button deleteCategoryButton = null!;
        private Button addSoftwareButton = null!;
        private Button deleteSoftwareButton = null!;
        private ComboBox schemeComboBox = null!;
        private Button saveSchemeButton = null!;
        private Button deleteSchemeButton = null!;

        private List<Control> bottomControls = new List<Control>(); // 声明为类成员

        private SoftwareManager softwareManager;
        private Dictionary<string, List<SoftwareItem>> savedSchemes = new Dictionary<string, List<SoftwareItem>>();
        private bool isUpdatingByCode = false;

        public MainForm()
        {
            InitializeComponent();
            softwareManager = new SoftwareManager();
            InitializeCustomComponents();
            InitializeSchemeManagement();
            PopulateSoftwareList();
        }

        private void InitializeCustomComponents()
        {
            // =================================================================
            // UI 现代化 (Edge 风格)
            // =================================================================

            this.BackColor = Color.White;
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            this.Text = "软件安装助理";
            this.Size = new System.Drawing.Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            // --- 顶部控件 ---
            categoryTreeView = new TreeView
            {
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(200, 380),
                BorderStyle = BorderStyle.FixedSingle,
                CheckBoxes = true,
                ShowLines = false // 更简洁的外观
            };

            softwareListView = new ListView
            {
                Location = new System.Drawing.Point(230, 20),
                Size = new System.Drawing.Size(540, 380),
                View = View.Details,
                BorderStyle = BorderStyle.FixedSingle,
                CheckBoxes = false,
                FullRowSelect = true
            };
            softwareListView.Columns.Add("序号", 50);
            softwareListView.Columns.Add("软件名称", 150);
            softwareListView.Columns.Add("版本", 80);
            softwareListView.Columns.Add("大小", 80);
            softwareListView.Columns.Add("说明", 180);

            addCategoryButton = new Button { Text = "➕ 新增分类", AutoSize = true, Location = new System.Drawing.Point(20, 410) };
            deleteCategoryButton = new Button { Text = "➖ 删除分类", AutoSize = true, Location = new System.Drawing.Point(125, 410) };
            addSoftwareButton = new Button { Text = "➕ 新增软件", AutoSize = true, Location = new System.Drawing.Point(20, 440) };
            deleteSoftwareButton = new Button { Text = "➖ 删除软件", AutoSize = true, Location = new System.Drawing.Point(125, 440) };

            var utilityButtons = new List<Button> { addCategoryButton, deleteCategoryButton, addSoftwareButton, deleteSoftwareButton };

            // --- 底部操作栏 ---
            Panel bottomPanel = new Panel
            {
                Size = new System.Drawing.Size(this.ClientSize.Width, 80),
                BackColor = Color.White,
                Dock = DockStyle.Bottom
            };
            bottomPanel.Paint += (s, e) => e.Graphics.DrawLine(new Pen(Color.Gainsboro, 1), 0, 0, bottomPanel.Width, 0);

            // --- 底部面板内的控件 ---
            // 使用固定大小初始化控件
            installPathTextBox = new TextBox { Size = new System.Drawing.Size(150, 23), Text = @"D:\Program Files" };
            browseButton = new Button { Text = "安装路径", Size = new System.Drawing.Size(80, 25) }; // 固定大小

            manualInstallButton = new Button { Text = "▶️ 手动安装", Size = new System.Drawing.Size(100, 25) }; // 固定大小
            autoInstallButton = new Button { Text = "⚡ 自动安装", Size = new System.Drawing.Size(100, 25) }; // 固定大小

            schemeComboBox = new ComboBox { Size = new System.Drawing.Size(120, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            saveSchemeButton = new Button { Text = "保存方案", Size = new System.Drawing.Size(80, 25) }; // 固定大小
            deleteSchemeButton = new Button { Text = "删除方案", Size = new System.Drawing.Size(80, 25) }; // 固定大小

            // 按所需顺序添加到 bottomControls 列表
            bottomControls.Add(installPathTextBox);
            bottomControls.Add(browseButton);
            bottomControls.Add(manualInstallButton);
            bottomControls.Add(autoInstallButton);
            bottomControls.Add(schemeComboBox);
            bottomControls.Add(saveSchemeButton);
            bottomControls.Add(deleteSchemeButton);

            // --- 样式和布局逻辑 ---
            Action<Button, bool> styleButton = (btn, isPrimary) =>
            {
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = 1;
                btn.FlatAppearance.BorderColor = Color.Gainsboro;
                btn.ForeColor = isPrimary ? Color.White : Color.Black;
                btn.BackColor = isPrimary ? Color.DodgerBlue : Color.FromArgb(240, 240, 240);
                
                var hoverColor = isPrimary ? Color.FromArgb(0, 102, 204) : Color.FromArgb(220, 220, 220);
                var originalColor = btn.BackColor;
                btn.MouseEnter += (s, e) => btn.BackColor = hoverColor;
                btn.MouseLeave += (s, e) => btn.BackColor = originalColor;
            };

            utilityButtons.ForEach(btn => styleButton(btn, false));
            styleButton(browseButton, false);
            styleButton(saveSchemeButton, false);
            styleButton(deleteSchemeButton, false);
            styleButton(manualInstallButton, true);
            styleButton(autoInstallButton, true);

            // 底部控件的动态布局
            this.Load += (s, e) => {
                int totalWidth = bottomControls.Sum(c => c.Width) + (bottomControls.Count - 1) * 10;
                int currentX = (bottomPanel.ClientSize.Width - totalWidth) / 2;
                int controlY = (bottomPanel.ClientSize.Height - manualInstallButton.Height) / 2;

                foreach (var control in bottomControls)
                {
                    int currentY = controlY + (manualInstallButton.Height - control.Height) / 2;
                    control.Location = new System.Drawing.Point(currentX, currentY);
                    bottomPanel.Controls.Add(control);
                    currentX += control.Width + 10;
                }
            };

            // 将所有控件添加到窗体
            this.Controls.Add(categoryTreeView);
            this.Controls.Add(softwareListView);
            this.Controls.Add(bottomPanel);
            utilityButtons.ForEach(btn => this.Controls.Add(btn));

            // --- 事件处理程序 ---
            categoryTreeView.AfterCheck += CategoryTreeView_AfterCheck;
            categoryTreeView.NodeMouseClick += CategoryTreeView_NodeMouseClick;
            addCategoryButton.Click += AddCategoryButton_Click;
            deleteCategoryButton.Click += DeleteCategoryButton_Click;
            addSoftwareButton.Click += AddSoftwareButton_Click;
            deleteSoftwareButton.Click += DeleteSoftwareButton_Click;
            browseButton.Click += BrowseButton_Click;
            manualInstallButton.Click += ManualInstallButton_Click;
            autoInstallButton.Click += AutoInstallButton_Click;
            schemeComboBox.SelectedIndexChanged += SchemeComboBox_SelectedIndexChanged;
            saveSchemeButton.Click += SaveSchemeButton_Click;
            deleteSchemeButton.Click += DeleteSchemeButton_Click;
        }

        private void InitializeSchemeManagement()
        {
            savedSchemes = SchemeManager.LoadSchemes();
            UpdateSchemeComboBox();
        }

        private void PopulateSoftwareList()
        {
            TreeNode allSoftwareNode = new TreeNode("所有软件");
            allSoftwareNode.Checked = false;
            categoryTreeView.Nodes.Add(allSoftwareNode);

            Dictionary<string, TreeNode> categoryNodes = new Dictionary<string, TreeNode>();

            Action<string> AddCategory = (name) =>
            {
                TreeNode node = new TreeNode(name);
                allSoftwareNode.Nodes.Add(node);
                categoryNodes.Add(name, node);
            };

            var categories = softwareManager.SoftwareItems.Select(i => i.Category).Distinct().ToList();
            categories.ForEach(c => AddCategory(c));

            foreach (var item in softwareManager.SoftwareItems)
            {
                if (categoryNodes.TryGetValue(item.Category, out TreeNode? categoryNode))
                {
                    TreeNode softwareTreeNode = new TreeNode(item.Name);
                    softwareTreeNode.Checked = item.IsSelected;
                    softwareTreeNode.Tag = item;
                    categoryNode.Nodes.Add(softwareTreeNode);
                }
            }

            categoryTreeView.ExpandAll();
        }

        private void CategoryTreeView_AfterCheck(object? sender, TreeViewEventArgs e)
        {
            if (isUpdatingByCode) return;

            isUpdatingByCode = true;
            if (e.Node != null)
            {
                if (e.Node.Tag is SoftwareItem clickedItem)
                {
                    clickedItem.IsSelected = e.Node.Checked;
                }
                UpdateChildNodeSelection(e.Node, e.Node.Checked);
                TreeNode? parent = e.Node.Parent;
                while (parent != null)
                {
                    UpdateParentNodeCheckState(parent);
                    parent = parent.Parent;
                }
                UpdateSoftwareListView();
            }
            isUpdatingByCode = false;
        }

        private void UpdateChildNodeSelection(TreeNode node, bool isChecked)
        {
            foreach (TreeNode childNode in node.Nodes)
            {
                childNode.Checked = isChecked;
                if (childNode.Tag is SoftwareItem childItem)
                {
                    childItem.IsSelected = isChecked;
                }
                if (childNode.Nodes.Count > 0)
                {
                    UpdateChildNodeSelection(childNode, isChecked);
                }
            }
        }

        private void UpdateParentNodeCheckState(TreeNode parentNode)
        {
            bool allChildrenChecked = parentNode.Nodes.Cast<TreeNode>().All(n => n.Checked);
            parentNode.Checked = allChildrenChecked;
        }

        private void UpdateSoftwareListView()
        {
            softwareListView.Items.Clear();
            int itemIndex = 1;

            void FindCheckedSoftware(TreeNodeCollection nodes)
            {
                foreach (TreeNode node in nodes)
                {
                    if (node.Tag is SoftwareItem item && node.Checked)
                    {
                        ListViewItem lvi = new ListViewItem(itemIndex.ToString());
                        lvi.SubItems.Add(item.Name);
                        lvi.SubItems.Add(item.Version);
                        lvi.SubItems.Add(item.Size);
                        lvi.SubItems.Add(item.Description);
                        lvi.Tag = item;
                        softwareListView.Items.Add(lvi);
                        itemIndex++;
                    }
                    if (node.Nodes.Count > 0)
                    {
                        FindCheckedSoftware(node.Nodes);
                    }
                }
            }
            FindCheckedSoftware(categoryTreeView.Nodes);
        }

        private void CategoryTreeView_NodeMouseClick(object? sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node == null) return; // 修复 CS8602：确保 e.Node 不为 null

            if (e.Button == MouseButtons.Left)
            {
                var hitTest = e.Node.TreeView?.HitTest(e.Location);
                if (hitTest != null && hitTest.Location == TreeViewHitTestLocations.Label)
                {
                    e.Node.Checked = !e.Node.Checked;
                }
            }
        }

        private void AddCategoryButton_Click(object? sender, EventArgs e)
        {
            using (var dialog = new InputDialog("增加分类", "请输入新的分类名称:"))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string? categoryName = dialog.InputText;
                    if (!string.IsNullOrWhiteSpace(categoryName))
                    {
                        TreeNode? allSoftwareNode = categoryTreeView.Nodes.Cast<TreeNode>().FirstOrDefault(n => n.Text == "所有软件");
                        if (allSoftwareNode != null)
                        {
                            if (allSoftwareNode.Nodes.Cast<TreeNode>().Any(n => n.Text == categoryName))
                            {
                                MessageBox.Show("该分类已存在！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            TreeNode newCategoryNode = new TreeNode(categoryName);
                            allSoftwareNode.Nodes.Add(newCategoryNode);
                            allSoftwareNode.Expand();
                        }
                    }
                }
            }
        }

        private void DeleteCategoryButton_Click(object? sender, EventArgs e)
        {
            TreeNode? selectedNode = categoryTreeView.SelectedNode;
            if (selectedNode != null && selectedNode.Parent != null && selectedNode.Parent.Text == "所有软件")
            {
                if (MessageBox.Show($"确定要删除分类 '{selectedNode.Text}' 吗?\n这将同时删除该分类下的所有软件。", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    softwareManager.SoftwareItems.RemoveAll(item => item.Category == selectedNode.Text);
                    selectedNode.Remove();
                    UpdateSoftwareListView();
                }
            }
            else
            {
                MessageBox.Show("请先选择一个要删除的分类.\n（注意：不能删除'所有软件'根分类）", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void AddSoftwareButton_Click(object? sender, EventArgs e)
        {
            TreeNode? selectedNode = categoryTreeView.SelectedNode;
            if (selectedNode == null || (selectedNode.Tag is SoftwareItem))
            {
                MessageBox.Show("请先在左侧选择一个分类，然后再添加软件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 防止在“所有软件”根节点下直接添加软件
            if (selectedNode.Text == "所有软件")
            {
                MessageBox.Show("不能直接在'所有软件'根分类下添加软件，请选择一个具体的分类。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "可执行文件 (*.exe)|*.exe|所有文件 (*.*)|*.*";
                openFileDialog.Title = "请选择软件安装包";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    var fileInfo = new FileInfo(filePath);
                    FileVersionInfo versionInfo = FileVersionInfo.GetVersionInfo(filePath);

                    string name = versionInfo.FileDescription ?? Path.GetFileNameWithoutExtension(fileInfo.Name);
                    string version = versionInfo.FileVersion ?? "N/A";
                    string size = $"{(fileInfo.Length / 1024.0 / 1024.0):F2} MB";
                    string? description = Microsoft.VisualBasic.Interaction.InputBox("请编辑软件说明:", "增加软件", versionInfo.Comments ?? "");
                    // TODO: 实现一个静默参数数据库（例如 a.json 文件），在添加软件时，根据文件名或描述自动查找并填充推荐的静默参数。
                    string? silentArgs = string.Empty;

                    var newSoftware = new SoftwareItem
                    {
                        Name = name, Version = version, Size = size, Description = description ?? "",
                        Category = selectedNode.Text, FilePath = filePath, SilentInstallArgs = silentArgs, IsSelected = true
                    };
                    softwareManager.SoftwareItems.Add(newSoftware);

                    TreeNode softwareNode = new TreeNode(newSoftware.Name) { Tag = newSoftware, Checked = true };
                    selectedNode.Nodes.Add(softwareNode);
                    selectedNode.Expand();
                    UpdateSoftwareListView();
                }
            }
        }

        private void DeleteSoftwareButton_Click(object? sender, EventArgs e)
        {
            TreeNode? selectedNode = categoryTreeView.SelectedNode;
            if (selectedNode != null && selectedNode.Tag is SoftwareItem itemToRemove)
            {
                if (MessageBox.Show($"确定要删除软件 '{selectedNode.Text}' 吗?", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    softwareManager.SoftwareItems.Remove(itemToRemove);
                    selectedNode.Remove();
                    UpdateSoftwareListView();
                }
            }
            else
            {
                MessageBox.Show("请先选择一个要删除的软件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BrowseButton_Click(object? sender, EventArgs e)
        {
            using (var folderBrowserDialog = new FolderBrowserDialog())
            {
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    installPathTextBox.Text = folderBrowserDialog.SelectedPath;
                }
            }
        }

        private void ManualInstallButton_Click(object? sender, EventArgs e)
        {
            List<SoftwareItem> selectedSoftware = GetSelectedSoftwareItems();
            if (selectedSoftware.Count == 0)
            {
                MessageBox.Show("没有选择任何软件进行安装。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var item in selectedSoftware)
            {
                ProcessRunner.ExecuteInstallation(item, false);
            }
        }

        private void AutoInstallButton_Click(object? sender, EventArgs e)
        {
            List<SoftwareItem> selectedSoftware = GetSelectedSoftwareItems();
            if (selectedSoftware.Count == 0)
            {
                MessageBox.Show("没有选择任何软件进行安装。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var item in selectedSoftware)
            {
                ProcessRunner.ExecuteInstallation(item, true);
            }
        }

        private List<SoftwareItem> GetSelectedSoftwareItems()
        {
            return softwareManager.SoftwareItems.Where(item => item.IsSelected).ToList();
        }

        private void SaveSchemeButton_Click(object? sender, EventArgs e)
        {
            using (var dialog = new InputDialog("保存方案", "请输入方案名称:"))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string? schemeName = dialog.InputText;
                    if (string.IsNullOrWhiteSpace(schemeName)) return;

                    if (savedSchemes.ContainsKey(schemeName))
                    {
                        if (MessageBox.Show("该方案名称已存在，要覆盖吗?", "确认覆盖", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                        {
                            return;
                        }
                    }

                    SyncModelFromTree(categoryTreeView.Nodes);
                    savedSchemes[schemeName] = new List<SoftwareItem>(softwareManager.SoftwareItems.Select(item => item.Clone()));
                    SchemeManager.SaveSchemes(savedSchemes);
                    UpdateSchemeComboBox();
                    schemeComboBox.SelectedItem = schemeName;
                }
            }
        }

        private void DeleteSchemeButton_Click(object? sender, EventArgs e)
        {
            if (schemeComboBox.SelectedItem is string selectedScheme && !string.IsNullOrEmpty(selectedScheme))
            {
                if (MessageBox.Show($"确定要删除方案 '{selectedScheme}' 吗?", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    savedSchemes.Remove(selectedScheme);
                    SchemeManager.SaveSchemes(savedSchemes);
                    UpdateSchemeComboBox();
                }
            }
            else
            {
                MessageBox.Show("请先从下拉框中选择一个要删除的方案。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void SchemeComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (schemeComboBox.SelectedItem is string selectedScheme && savedSchemes.TryGetValue(selectedScheme, out var loadedItems))
            {
                var selectedInScheme = new HashSet<string>(loadedItems.Where(item => item.IsSelected).Select(item => item.Name));
                foreach (var item in softwareManager.SoftwareItems)
                {
                    item.IsSelected = selectedInScheme.Contains(item.Name);
                }

                isUpdatingByCode = true;
                RebuildTreeView();
                isUpdatingByCode = false;

                UpdateSoftwareListView();
            }
        }

        private void SyncModelFromTree(TreeNodeCollection nodes)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Tag is SoftwareItem item)
                {
                    item.IsSelected = node.Checked;
                }
                if (node.Nodes.Count > 0)
                {
                    SyncModelFromTree(node.Nodes);
                }
            }
        }

        private void UpdateSchemeComboBox()
        {
            var currentSelection = schemeComboBox.SelectedItem;
            schemeComboBox.Items.Clear();
            foreach (var schemeName in savedSchemes.Keys.OrderBy(k => k))
            {
                schemeComboBox.Items.Add(schemeName);
            }
            if (currentSelection != null && schemeComboBox.Items.Contains(currentSelection))
            {
                schemeComboBox.SelectedItem = currentSelection;
            }
        }

        private void RebuildTreeView()
        {
            categoryTreeView.Nodes.Clear();
            TreeNode allSoftwareNode = new TreeNode("所有软件");
            categoryTreeView.Nodes.Add(allSoftwareNode);

            var categoryNodes = softwareManager.SoftwareItems.Select(i => i.Category).Distinct()
                .ToDictionary(name => name, name => new TreeNode(name));

            foreach (var node in categoryNodes.Values.OrderBy(n => n.Text))
            {
                allSoftwareNode.Nodes.Add(node);
            }

            foreach (var item in softwareManager.SoftwareItems)
            {
                if (categoryNodes.TryGetValue(item.Category, out var categoryNode))
                {
                    TreeNode softwareNode = new TreeNode(item.Name) { Tag = item, Checked = item.IsSelected };
                    categoryNode.Nodes.Add(softwareNode);
                }
            }

            foreach(TreeNode categoryNode in allSoftwareNode.Nodes)
            {
                UpdateParentNodeCheckState(categoryNode);
            }
            UpdateParentNodeCheckState(allSoftwareNode);

            categoryTreeView.ExpandAll();
        }
    }
}
