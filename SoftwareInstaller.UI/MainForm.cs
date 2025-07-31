using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using SoftwareInstaller.Core;
using SoftwareInstaller.Models;
using SoftwareInstaller.Utils;

namespace SoftwareInstaller.UI
{
    public partial class MainForm : Form
    {
        private readonly List<Control> _bottomControls = new List<Control>(); // 声明为类成员
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
                // Use the existing global error handler utility
                ErrorHandler.ShowError($"应用程序初始化失败，即将退出。\n错误: {ex.Message}");
                // Schedule the form to close immediately after it's loaded and shown.
                this.Load += (s, e) => this.Close();
                return;
            }

            InitializeCustomComponents(); // _softwareManager and _schemeHandler are not null here
            InitializeSchemeManagement(); // _schemeHandler is not null here
            PopulateSoftwareList();       // _softwareManager is not null here
        }

        private void InitializeCustomComponents()
        {
            // =================================================================
            // UI 现代化 (Edge 风格)
            // =================================================================
            // 控件现在由设计器创建。这里只处理不能在设计器中完成的逻辑。
            
            // --- 底部面板内的控件 ---
            // 这些控件在设计器中创建，但我们需要在这里实例化它们以便进行动态布局和样式设置
            installPathTextBox = new TextBox { Size = new System.Drawing.Size(150, 23), Text = DefaultInstallPath };
            browseButton = new Button { Text = "安装路径", Size = new System.Drawing.Size(80, 25) };
            manualInstallButton = new Button { Text = "▶️ 手动安装", Size = new System.Drawing.Size(100, 25) };
            autoInstallButton = new Button { Text = "⚡ 自动安装", Size = new System.Drawing.Size(100, 25) };
            schemeComboBox = new ComboBox { Size = new System.Drawing.Size(120, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            saveSchemeButton = new Button { Text = "保存方案", Size = new System.Drawing.Size(80, 25) };
            deleteSchemeButton = new Button { Text = "删除方案", Size = new System.Drawing.Size(80, 25) };
            
            var utilityButtons = new List<Button> { addCategoryButton, deleteCategoryButton, addSoftwareButton, deleteSoftwareButton };

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

            // 将底部控件添加到列表中以便动态布局
            _bottomControls.AddRange(new Control[] { installPathTextBox, browseButton, manualInstallButton, autoInstallButton, schemeComboBox, saveSchemeButton, deleteSchemeButton });

            // --- 事件处理程序 ---
            this.Load += OnMainFormLoad;
            bottomPanel.Paint += BottomPanel_Paint;
            bottomPanel.Resize += OnBottomPanelResize;

            // 控件事件
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

            // --- 自动保存事件 ---
            this.FormClosing += MainForm_FormClosing;
        }

        private void OnMainFormLoad(object? sender, EventArgs e)
        {
            // 首次加载时，将控件添加到面板
            foreach (var control in _bottomControls)
            {
                bottomPanel.Controls.Add(control);
            }
            // 并执行初始布局
            LayoutBottomControls();
        }

        private void OnBottomPanelResize(object? sender, EventArgs e)
        {
            // 每当面板大小变化时，重新布局
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
            // 绘制顶部边框线
            e.Graphics.DrawLine(new Pen(Color.Gainsboro, 1), 0, 0, bottomPanel.Width, 0);
        }

        private void InitializeSchemeManagement()
        {
            PopulateSchemeComboBox();
        }

        private void PopulateSoftwareList()
        {
            TreeNode allSoftwareNode = new TreeNode(AllSoftwareNodeText);
            allSoftwareNode.Checked = false;
            categoryTreeView.Nodes.Add(allSoftwareNode);

            Dictionary<string, TreeNode> categoryNodes = new Dictionary<string, TreeNode>();

            Action<string> AddCategory = (name) =>
            {
                TreeNode node = new TreeNode(name);
                allSoftwareNode.Nodes.Add(node);
                categoryNodes.Add(name, node);
            };

            var categories = _softwareManager!.SoftwareItems.Select(i => i.Category).Distinct().ToList();
            categories.ForEach(c => AddCategory(c));

            foreach (var item in _softwareManager!.SoftwareItems)
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
            // 当通过代码更新UI时（例如加载方案时），isUpdatingByCode会为true，此时我们跳过此事件处理，防止不必要的操作。
            if (isUpdatingByCode || e.Node == null) return;

            // 1. 临时分离事件处理器，防止后续的编程修改触发连锁反应。这是关键的健壮性改进。
            categoryTreeView.AfterCheck -= CategoryTreeView_AfterCheck;

            try
            {
                // 2. 更新UI：将当前操作节点的勾选状态向下传递给所有子节点。
                // This check is now here to prevent exceptions if e.Node is null
                SetChildNodeCheckedState(e.Node, e.Node.Checked);
                

                // 3. 更新UI：从当前节点开始，向上更新所有父节点的勾选状态。
                // e.Node is already checked for null at the beginning of the method.
                TreeNode? parent = e.Node.Parent;
                while (parent != null)
                {
                    UpdateParentNodeCheckState(parent);
                    parent = parent.Parent;
                }

                // 4. 在所有UI更新完毕后，将TreeView的最终状态一次性同步到数据模型中。
                SyncModelFromTree(categoryTreeView.Nodes);

                // 5. 根据更新后的数据模型，刷新右侧的ListView。
                UpdateSoftwareListView();
            }
            finally
            {
                // 6. 无论成功与否，都必须将事件处理器重新附加回去。
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

        private void CategoryTreeView_NodeMouseDoubleClick(object? sender, TreeNodeMouseClickEventArgs e)
        {
            // 我们只关心代表软件的节点，这些节点的 Tag 属性是 SoftwareItem 类型
            if (e.Node?.Tag is SoftwareItem itemToEdit)
            {
                // 使用现有的软件条目数据打开编辑窗体
                using (var form = new SoftwareEditForm(itemToEdit))
                {
                    if (form.ShowDialog(this) == DialogResult.OK)
                    {
                        // 当用户点击“确定”后，传入的 'itemToEdit' 对象已经被窗体修改。
                        // 我们只需要刷新UI来反映这些更改。

                        // 由于软件名称可能已更改，最安全的方式是完全重建UI。
                        // 这可以确保排序和列表视图都保持同步。
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
            TreeNode? selectedNode = categoryTreeView.SelectedNode;
            if (selectedNode != null && selectedNode.Parent != null && selectedNode.Parent.Text == AllSoftwareNodeText)
            {
                if (MessageBox.Show($"确定要删除分类 '{selectedNode.Text}' 吗?\n这将同时删除该分类下的所有软件。", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    _softwareManager!.SoftwareItems.RemoveAll(item => item.Category == selectedNode.Text);
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
            if (selectedNode.Text == AllSoftwareNodeText)
            {
                MessageBox.Show("不能直接在'所有软件'根分类下添加软件，请选择一个具体的分类。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 1. 先让用户选择文件
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "可执行文件 (*.exe)|*.exe|所有文件 (*.*)|*.*";
                openFileDialog.Title = "请选择软件安装包";

                if (openFileDialog.ShowDialog(this) == DialogResult.OK)
                {
                    // 2. 根据文件自动填充信息
                    var newSoftware = new SoftwareItem
                    {
                        FilePath = openFileDialog.FileName,
                        Name = string.Empty,
                        Version = string.Empty,
                        Size = string.Empty,
                        Description = string.Empty,
                        Category = string.Empty // This will be overwritten later
                    };
                    try
                    {
                        var fileInfo = new FileInfo(newSoftware.FilePath);
                        var versionInfo = FileVersionInfo.GetVersionInfo(newSoftware.FilePath);

                        newSoftware.Name = versionInfo.FileDescription ?? Path.GetFileNameWithoutExtension(fileInfo.Name);
                        newSoftware.Version = versionInfo.FileVersion ?? "N/A";
                        newSoftware.Size = $"{(fileInfo.Length / 1024.0 / 1024.0):F2} MB";
                    }
                    catch (Exception ex)
                    {
                        // 如果无法读取文件信息，这也不是致命错误，程序可以继续
                        Debug.WriteLine($"无法读取文件信息: {ex.Message}");
                        // 至少用文件名填充
                        newSoftware.Name = Path.GetFileNameWithoutExtension(newSoftware.FilePath);
                    }

                    // 3. 打开编辑/确认窗口，并预填好数据
                    using (var form = new SoftwareEditForm(newSoftware))
                    {
                        form.Text = "新增软件"; // 明确设置为“新增”模式
                        if (form.ShowDialog(this) == DialogResult.OK)
                        {
                            // 用户确认后，SoftwareItem 已经被更新
                            var finalSoftware = form.SoftwareItem;
                            finalSoftware.Category = selectedNode.Text;
                            finalSoftware.IsSelected = true;
                            _softwareManager!.SoftwareItems.Add(finalSoftware);

                            RebuildTreeView();
                            UpdateSoftwareListView();
                        }
                    }
                }
            }
        }

        private void DeleteSoftwareButton_Click(object? sender, EventArgs e)
        {
            List<SoftwareItem> itemsToRemove = GetSelectedSoftwareItems();

            if (itemsToRemove.Count == 0)
            {
                MessageBox.Show("请先勾选需要删除的软件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var confirmResult = MessageBox.Show($"您确定要删除选中的 {itemsToRemove.Count} 款软件吗？\n此操作不可恢复。", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (confirmResult == DialogResult.Yes)
            {
                // 从数据模型中移除
                foreach (var item in itemsToRemove)
                {
                    _softwareManager!.SoftwareItems.Remove(item);
                }

                // 从UI (TreeView) 中精准移除对应的节点，而不是完全重建
                // 这样可以保留空的分类
                List<TreeNode> nodesToRemove = new List<TreeNode>();
                FindNodesBySoftwareItems(categoryTreeView.Nodes, itemsToRemove, nodesToRemove);

                foreach (var node in nodesToRemove)
                {
                    var parent = node.Parent;
                    node.Remove();
                    // 移除软件节点后，更新其父分类节点的勾选状态
                    if (parent != null)
                    {
                        UpdateParentNodeCheckState(parent);
                    }
                }

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
            return _softwareManager!.SoftwareItems.Where(item => item.IsSelected).ToList();
        }

        private void SaveSchemeButton_Click(object? sender, EventArgs e)
        {
            using (var dialog = new InputDialog("保存方案", "请输入方案名称:"))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string? schemeName = dialog.InputText;
                    if (string.IsNullOrWhiteSpace(schemeName)) return;

                    if (_schemeHandler!.SchemeExists(schemeName))
                    {
                        if (MessageBox.Show("该方案名称已存在，要覆盖吗?", "确认覆盖", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                        {
                            return;
                        }
                    }

                    _schemeHandler!.SaveScheme(schemeName, _softwareManager!.SoftwareItems);
                    PopulateSchemeComboBox();
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
                    _schemeHandler!.DeleteScheme(selectedScheme);
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
            if (schemeComboBox.SelectedItem is string selectedScheme && _schemeHandler!.GetScheme(selectedScheme) is List<SoftwareItem> loadedItems)
            {
                var selectedInScheme = new HashSet<string>(loadedItems.Where(item => item.IsSelected).Select(item => item.Name));
                foreach (var item in _softwareManager!.SoftwareItems)
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

        private void PopulateSchemeComboBox()
        {
            var currentSelection = schemeComboBox.SelectedItem;
            schemeComboBox.Items.Clear();
            foreach (var schemeName in _schemeHandler!.GetSchemeNames())
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

            var categoryNodes = _softwareManager!.SoftwareItems.Select(i => i.Category).Distinct()
                .ToDictionary(name => name, name => new TreeNode(name));

            foreach (var node in categoryNodes.Values.OrderBy(n => n.Text))
            {
                allSoftwareNode.Nodes.Add(node);
            }
            
            foreach (var item in _softwareManager!.SoftwareItems)
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

        private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
        {
            // 在窗口关闭前，自动保存对软件列表的所有更改。
            _softwareManager!.SaveSoftwareList();
        }
    }
}