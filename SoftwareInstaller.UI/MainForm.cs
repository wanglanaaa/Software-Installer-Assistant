using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using SoftwareInstaller.Core;
using SoftwareInstaller.Models;
using SoftwareInstaller.Utils;

namespace SoftwareInstaller.UI
{
    public partial class MainForm : Form
    {
        private readonly List<Control> _bottomControls = new List<Control>();
        private readonly SoftwareManager? _softwareManager;
        private readonly SchemeHandler? _schemeHandler;
        private const string AllSoftwareNodeText = "所有软件";
        private const string DefaultInstallPath = @"D:\Program Files";
        private bool isUpdatingByCode = false;

        public MainForm()
        {
            InitializeComponent();

            try
            {
                _softwareManager = new SoftwareManager();
                _schemeHandler = new SchemeHandler();
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowError($"应用程序初始化失败，即将退出。\n错误: {ex.Message}");
                this.Load += (s, e) => this.Close();
                return;
            }

            InitializeCustomComponents();
            InitializeSchemeManagement();
            RebuildTreeView(); 
        }

        private void InitializeCustomComponents()
        {
            installPathTextBox = new TextBox { Size = new System.Drawing.Size(150, 23), Text = DefaultInstallPath };
            browseButton = new Button { Text = "安装路径", Size = new System.Drawing.Size(80, 25) };
            manualInstallButton = new Button { Text = "▶️ 手动安装", Size = new System.Drawing.Size(100, 25) };
            autoInstallButton = new Button { Text = "⚡ 自动安装", Size = new System.Drawing.Size(100, 25) };
            schemeComboBox = new ComboBox { Size = new System.Drawing.Size(120, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            saveSchemeButton = new Button { Text = "保存方案", Size = new System.Drawing.Size(80, 25) };
            deleteSchemeButton = new Button { Text = "删除方案", Size = new System.Drawing.Size(80, 25) };
            
            var utilityButtons = new List<Button> { addCategoryButton, deleteCategoryButton, addSoftwareButton, deleteSoftwareButton };

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

            _bottomControls.AddRange(new Control[] { installPathTextBox, browseButton, manualInstallButton, autoInstallButton, schemeComboBox, saveSchemeButton, deleteSchemeButton });

            this.Load += OnMainFormLoad;
            bottomPanel.Paint += BottomPanel_Paint;
            bottomPanel.Resize += OnBottomPanelResize;

            categoryTreeView.AfterCheck += CategoryTreeView_AfterCheck;
            categoryTreeView.NodeMouseClick += CategoryTreeView_NodeMouseClick;
            categoryTreeView.NodeMouseDoubleClick += CategoryTreeView_NodeMouseDoubleClick;
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

            this.FormClosing += MainForm_FormClosing;
        }

        private void OnMainFormLoad(object? sender, EventArgs e)
        {
            foreach (var control in _bottomControls)
            {
                bottomPanel.Controls.Add(control);
            }
            LayoutBottomControls();
        }

        private void OnBottomPanelResize(object? sender, EventArgs e)
        {
            LayoutBottomControls();
        }

        private void LayoutBottomControls()
        {
            if (!_bottomControls.Any()) return;

            int totalWidth = _bottomControls.Sum(c => c.Width) + (_bottomControls.Count - 1) * 10;
            int currentX = (bottomPanel.ClientSize.Width - totalWidth) / 2;
            int controlY = (bottomPanel.ClientSize.Height - manualInstallButton.Height) / 2;

            foreach (var control in _bottomControls)
            {
                int currentY = controlY + (manualInstallButton.Height - control.Height) / 2;
                control.Location = new System.Drawing.Point(currentX, currentY);
                currentX += control.Width + 10;
            }
        }

        private void BottomPanel_Paint(object? sender, PaintEventArgs e)
        {
            e.Graphics.DrawLine(new Pen(Color.Gainsboro, 1), 0, 0, bottomPanel.Width, 0);
        }

        private void InitializeSchemeManagement()
        {
            PopulateSchemeComboBox();
        }

        private void RebuildTreeView()
        {
            categoryTreeView.Nodes.Clear();
            TreeNode allSoftwareNode = new TreeNode(AllSoftwareNodeText);
            categoryTreeView.Nodes.Add(allSoftwareNode);

            if (_softwareManager == null) return;

            var categoryNodes = _softwareManager.ActiveSoftwareItems.Select(i => i.Category).Distinct()
                .ToDictionary(name => name, name => new TreeNode(name));

            foreach (var node in categoryNodes.Values.OrderBy(n => n.Text))
            {
                allSoftwareNode.Nodes.Add(node);
            }
            
            foreach (var item in _softwareManager.ActiveSoftwareItems)
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

        private void CategoryTreeView_AfterCheck(object? sender, TreeViewEventArgs e)
        {
            if (isUpdatingByCode || e.Node == null) return;

            categoryTreeView.AfterCheck -= CategoryTreeView_AfterCheck;

            try
            {
                SetChildNodeCheckedState(e.Node, e.Node.Checked);
                
                TreeNode? parent = e.Node.Parent;
                while (parent != null)
                {
                    UpdateParentNodeCheckState(parent);
                    parent = parent.Parent;
                }

                SyncModelFromTree(categoryTreeView.Nodes);
                UpdateSoftwareListView();
            }
            finally
            {
                categoryTreeView.AfterCheck += CategoryTreeView_AfterCheck;
            }
        }

        private void SetChildNodeCheckedState(TreeNode node, bool isChecked)
        {
            foreach (TreeNode childNode in node.Nodes)
            {
                childNode.Checked = isChecked;
                if (childNode.Nodes.Count > 0)
                {
                    SetChildNodeCheckedState(childNode, isChecked);
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
            if (e.Node == null) return;

            if (e.Button == MouseButtons.Left)
            {
                var hitTest = e.Node.TreeView?.HitTest(e.Location);
                if (hitTest != null && hitTest.Location == TreeViewHitTestLocations.Label)
                {
                    e.Node.Checked = !e.Node.Checked;
                }
            }
        }

        private void CategoryTreeView_NodeMouseDoubleClick(object? sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node?.Tag is SoftwareItem itemToEdit)
            {
                using (var form = new SoftwareEditForm(itemToEdit))
                {
                    if (form.ShowDialog(this) == DialogResult.OK)
                    {
                        RebuildTreeView();
                        UpdateSoftwareListView();
                    }
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
                    if (!string.IsNullOrWhiteSpace(categoryName) && categoryName != AllSoftwareNodeText)
                    {
                        TreeNode? allSoftwareNode = categoryTreeView.Nodes.Cast<TreeNode>().FirstOrDefault(n => n.Text == AllSoftwareNodeText);
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
            if (_softwareManager == null) return;
            TreeNode? selectedNode = categoryTreeView.SelectedNode;
            if (selectedNode != null && selectedNode.Parent != null && selectedNode.Parent.Text == AllSoftwareNodeText)
            {
                if (MessageBox.Show($"确定要删除分类 '{selectedNode.Text}' 吗?\n这将同时删除该分类下的所有软件。", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    _softwareManager.ActiveSoftwareItems.RemoveAll(item => item.Category == selectedNode.Text);
                    RebuildTreeView();
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
            if (_softwareManager == null) return;
            TreeNode? selectedNode = categoryTreeView.SelectedNode;
            string categoryName;

            // 校验：确保用户没有选择根节点
            if (selectedNode != null && selectedNode.Text == AllSoftwareNodeText)
            {
                MessageBox.Show("不能直接在'所有软件'根分类下添加软件，请选择一个具体的分类，或创建新分类。", "操作无效", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (selectedNode == null)
            {
                categoryName = "未分类";
            }
            else if (selectedNode.Tag is SoftwareItem item)
            {
                categoryName = item.Category;
            }
            else
            {
                categoryName = selectedNode.Text;
            }

            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "可执行文件 (*.exe)|*.exe|所有文件 (*.*)|*.*";
                openFileDialog.Title = "请选择软件安装包";

                if (openFileDialog.ShowDialog(this) == DialogResult.OK)
                {
                    var newSoftware = new SoftwareItem
                    {
                        FilePath = openFileDialog.FileName,
                        Name = string.Empty,
                        Version = string.Empty,
                        Size = string.Empty,
                        Description = string.Empty,
                        Category = categoryName
                    };

                    try
                    {
                        var fileInfo = new FileInfo(newSoftware.FilePath);
                        var versionInfo = FileVersionInfo.GetVersionInfo(newSoftware.FilePath);
                        newSoftware.Name = (versionInfo.ProductName ?? Path.GetFileNameWithoutExtension(fileInfo.Name)).Trim();
                        newSoftware.Version = (versionInfo.FileVersion ?? "N/A").Trim();
                        newSoftware.Size = $"{(fileInfo.Length / 1024.0 / 1024.0):F2} MB";
                        newSoftware.Description = (versionInfo.FileDescription ?? string.Empty).Trim();
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"无法读取文件信息: {ex.Message}");
                        newSoftware.Name = Path.GetFileNameWithoutExtension(newSoftware.FilePath);
                    }

                    using (var form = new SoftwareEditForm(newSoftware))
                    {
                        form.Text = "新增软件";
                        if (form.ShowDialog(this) == DialogResult.OK)
                        {
                            var finalSoftware = form.SoftwareItem;
                            finalSoftware.Category = categoryName;
                            finalSoftware.IsSelected = true;

                            _softwareManager.AddSoftwareToActiveList(finalSoftware);

                            RebuildTreeView();
                            UpdateSoftwareListView();
                        }
                    }
                }
            }
        }

        private void DeleteSoftwareButton_Click(object? sender, EventArgs e)
        {
            if (_softwareManager == null) return;
            List<SoftwareItem> itemsToRemove = GetSelectedSoftwareItems();

            if (itemsToRemove.Count == 0)
            {
                MessageBox.Show("请先勾选需要删除的软件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string message = $"您确定要从当前列表中移除选中的 {itemsToRemove.Count} 款软件吗?\n(注意：这仅影响当前视图，不会从您的永久配置中删除它)";
            var confirmResult = MessageBox.Show(message, "确认移除", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (confirmResult == DialogResult.Yes)
            {
                foreach (var item in itemsToRemove)
                {
                    _softwareManager.ActiveSoftwareItems.Remove(item);
                }

                RebuildTreeView();
                UpdateSoftwareListView();
            }
        }

        private void FindNodesBySoftwareItems(TreeNodeCollection nodes, List<SoftwareItem> itemsToFind, List<TreeNode> foundNodes)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Tag is SoftwareItem item && itemsToFind.Contains(item))
                {
                    foundNodes.Add(node);
                }

                if (node.Nodes.Count > 0)
                {
                    FindNodesBySoftwareItems(node.Nodes, itemsToFind, foundNodes);
                }
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

        private async void AutoInstallButton_Click(object? sender, EventArgs e)
        {
            if (_softwareManager == null) return;
            List<SoftwareItem> selectedSoftware = GetSelectedSoftwareItems();
            if (selectedSoftware.Count == 0)
            {
                MessageBox.Show("没有选择任何软件进行安装。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            autoInstallButton.Enabled = false;
            this.UseWaitCursor = true;

            foreach (var item in selectedSoftware)
            {
                if (string.IsNullOrEmpty(item.FilePath) || !File.Exists(item.FilePath))
                {
                    MessageBox.Show($"软件 '{item.Name}' 的安装文件路径无效或不存在，已跳过。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue;
                }

                Func<Task<bool>> confirmInstallation = async () =>
                {
                    return await Task.Run(() =>
                    {
                        string message = $"安装程序 '{item.Name}' 的一个静默模式似乎已执行完毕。\n\n请您确认软件是否已成功安装到系统中？\n\n- 点击 '是'，程序将保存当前有效的静默参数。\n- 点击 '否'，程序将尝试下一个静默参数。";
                        var result = MessageBox.Show(message,
                            "请确认安装结果",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);
                        return result == DialogResult.Yes;
                    });
                };

                var (success, usedArgs) = await _softwareManager.InstallSoftwareIntelligently(item, confirmInstallation);

                if (success)
                {
                    string successMessage = $"软件 '{item.Name}' 已成功安装。\n使用的静默参数是: {usedArgs}";
                    MessageBox.Show(successMessage, "安装成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    string failMessage = $"无法为软件 '{item.Name}' 找到有效的静默安装参数。\n所有尝试都失败了，或者您取消了确认。\n建议您使用“手动安装”。";
                    MessageBox.Show(failMessage, "自动安装失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            autoInstallButton.Enabled = true;
            this.UseWaitCursor = false;
        }

        private List<SoftwareItem> GetSelectedSoftwareItems()
        {
            if (_softwareManager == null) return new List<SoftwareItem>();
            return _softwareManager.ActiveSoftwareItems.Where(item => item.IsSelected).ToList();
        }

        private void SaveSchemeButton_Click(object? sender, EventArgs e)
        {
            if (_schemeHandler == null || _softwareManager == null) return;
            using (var dialog = new InputDialog("保存方案", "请输入方案名称:"))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string? schemeName = dialog.InputText;
                    if (string.IsNullOrWhiteSpace(schemeName)) return;

                    if (_schemeHandler.SchemeExists(schemeName))
                    {
                        if (MessageBox.Show("该方案名称已存在，要覆盖吗?", "确认覆盖", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                        {
                            return;
                        }
                    }

                    _schemeHandler.SaveScheme(schemeName, _softwareManager.ActiveSoftwareItems);
                    PopulateSchemeComboBox();
                    schemeComboBox.SelectedItem = schemeName;
                }
            }
        }

        private void DeleteSchemeButton_Click(object? sender, EventArgs e)
        {
            if (_schemeHandler == null) return;
            if (schemeComboBox.SelectedItem is string selectedScheme && !string.IsNullOrEmpty(selectedScheme))
            {
                if (MessageBox.Show($"确定要删除方案 '{selectedScheme}' 吗?", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    _schemeHandler.DeleteScheme(selectedScheme);
                    PopulateSchemeComboBox();
                }
            }
            else
            {
                MessageBox.Show("请先从下拉框中选择一个要删除的方案。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void SchemeComboBox_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (_schemeHandler == null || _softwareManager == null) return;
            if (schemeComboBox.SelectedItem is string selectedScheme && _schemeHandler.GetScheme(selectedScheme) is List<SoftwareItem> loadedItems)
            {
                _softwareManager.ActiveSoftwareItems.Clear();
                foreach (var item in loadedItems)
                {
                    _softwareManager.AddSoftwareToActiveList(item);
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

        private void PopulateSchemeComboBox()
        {
            if (_schemeHandler == null) return;
            var currentSelection = schemeComboBox.SelectedItem;
            schemeComboBox.Items.Clear();
            foreach (var schemeName in _schemeHandler.GetSchemeNames())
            {
                schemeComboBox.Items.Add(schemeName);
            }
            if (currentSelection != null && schemeComboBox.Items.Contains(currentSelection))
            {
                schemeComboBox.SelectedItem = currentSelection;
            }
        }

        private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
        {
            _softwareManager?.SaveSoftwareList();
        }
    }
}
