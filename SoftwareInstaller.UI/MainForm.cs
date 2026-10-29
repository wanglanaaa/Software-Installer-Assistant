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
using SoftwareInstaller.Packaging;

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
        private readonly Button buildBundleButton = new() { Text = "生成单文件安装包", Size = new(155, 32) };
        private readonly Label selectionSummary = new() { AutoSize = true };
        private const int ConfigurationBarHeight = 42;
        private readonly Label configurationPathLabel = new() { Dock = DockStyle.Fill, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft };
        private readonly ToolTip configurationTips = new();
        private string? configurationPath;
        private string savedConfiguration = ConfigurationStore.Serialize(new WorkspaceConfiguration());

        public MainForm()
        {
            InitializeComponent();
            AppBranding.Apply(this);

            try
            {
                _softwareManager = new SoftwareManager(useDefaultStorage: false);
                _schemeHandler = new SchemeHandler(useDefaultStorage: false);
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowError($"应用程序初始化失败，即将退出。\n错误: {ex.Message}");
                this.Load += (s, e) => this.Close();
                return;
            }

            InitializeCustomComponents();
            InitializeConfigurationBar();
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
            styleButton(buildBundleButton, true);
            buildBundleButton.Click += (_, _) =>
            {
                if (_softwareManager == null || _schemeHandler == null) return;
                using var dialog = new BundleBuildForm(_softwareManager.ActiveSoftwareItems, _schemeHandler, _softwareManager.ActiveCategories);
                dialog.ShowDialog(this);
            };
            bottomPanel.Height = 100;
            bottomPanel.Dock = DockStyle.Bottom;
            MinimumSize = new Size(820, 650);
            // A global path is not supported by arbitrary third-party installers.
            installPathTextBox.Visible = browseButton.Visible = false;

            _bottomControls.AddRange(new Control[] { schemeComboBox, saveSchemeButton, deleteSchemeButton, selectionSummary, manualInstallButton, autoInstallButton, buildBundleButton });

            this.Load += OnMainFormLoad;
            this.Resize += (_, _) => LayoutContent();
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
            LayoutContent();
        }

        private void OnBottomPanelResize(object? sender, EventArgs e)
        {
            LayoutBottomControls();
        }

        private void LayoutContent()
        {
            int top = ConfigurationBarHeight;
            int bodyHeight = Math.Max(220, ClientSize.Height - bottomPanel.Height - top);
            softwareListView.SetBounds(230, top + 20, Math.Max(300, ClientSize.Width - 250), bodyHeight - 40);
            categoryTreeView.SetBounds(20, top + 20, 200, bodyHeight - 112);
            addCategoryButton.Location = new Point(20, top + bodyHeight - 82);
            deleteCategoryButton.Location = new Point(124, top + bodyHeight - 82);
            addSoftwareButton.Location = new Point(20, top + bodyHeight - 48);
            deleteSoftwareButton.Location = new Point(124, top + bodyHeight - 48);
        }

        private void InitializeConfigurationBar()
        {
            var bar = new TableLayoutPanel { Dock = DockStyle.Top, Height = ConfigurationBarHeight, ColumnCount = 6, RowCount = 1, Padding = new(16, 6, 16, 4) };
            bar.RowStyles.Add(new(SizeType.Percent, 100));
            bar.ColumnStyles.Add(new(SizeType.Absolute, 40));
            var logo = new PictureBox { Image = AppBranding.Image, SizeMode = PictureBoxSizeMode.Zoom, Dock = DockStyle.Fill, Margin = new(0, 0, 8, 0) };
            bar.Controls.Add(logo, 0, 0); configurationTips.SetToolTip(logo, "软件安装助理");
            for (int i = 0; i < 4; i++) bar.ColumnStyles.Add(new(SizeType.Absolute, 90));
            bar.ColumnStyles.Add(new(SizeType.Percent, 100));
            void AddButton(string text, int column, Action action)
            {
                var button = new Button { Text = text, Dock = DockStyle.Fill, Margin = new(0, 0, 6, 0) };
                button.Click += (_, _) => action(); bar.Controls.Add(button, column + 1, 0);
            }
            AddButton("新建配置", 0, () => { if (ConfirmConfigurationChange()) ApplyConfiguration(new(), null); });
            AddButton("加载配置…", 1, () =>
            {
                using var dialog = new OpenFileDialog { Filter = "安装配置 (*.json)|*.json", Title = "加载软件、分类和安装方案配置" };
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try { LoadConfiguration(dialog.FileName); }
                catch (Exception ex) { ErrorHandler.ShowError("加载配置失败，当前内容保持不变：" + ex.Message); }
            });
            AddButton("保存配置", 2, () => SaveConfiguration(false));
            AddButton("另存为…", 3, () => SaveConfiguration(true));
            bar.Controls.Add(configurationPathLabel, 5, 0); Controls.Add(bar);
            configurationTips.SetToolTip(saveSchemeButton, "将当前列表和勾选状态记为方案；点击上方「保存配置」写入配置文件。");
            FormClosed += (_, _) => configurationTips.Dispose();
            UpdateConfigurationPath();
        }

        private WorkspaceConfiguration CaptureConfiguration() => new()
        {
            Software = _softwareManager!.ActiveSoftwareItems.Select(s => s.Clone()).ToList(),
            Categories = new(_softwareManager.ActiveCategories),
            Schemes = _schemeHandler!.GetDefinitions()
        };

        private bool LoadConfiguration(string path)
        {
            var configuration = ConfigurationStore.Load(path, out bool importedLegacy);
            if (!ConfirmConfigurationChange()) return false;
            // The confirmation may save this same file; load the saved version rather than the earlier snapshot.
            configuration = ConfigurationStore.Load(path, out importedLegacy);
            ApplyConfiguration(configuration, importedLegacy ? null : Path.GetFullPath(path));
            if (importedLegacy) savedConfiguration = ConfigurationStore.Serialize(new WorkspaceConfiguration());
            return true;
        }

        private void ApplyConfiguration(WorkspaceConfiguration configuration, string? path)
        {
            schemeComboBox.SelectedIndex = -1;
            _softwareManager!.ReplaceActiveList(configuration.Software, configuration.Categories);
            _schemeHandler!.ReplaceSchemes(configuration.Schemes);
            configurationPath = path;
            PopulateSchemeComboBox(); RebuildTreeView();
            savedConfiguration = ConfigurationStore.Serialize(CaptureConfiguration());
            UpdateConfigurationPath();
        }

        private void UpdateConfigurationPath()
        {
            configurationPathLabel.Text = configurationPath == null ? "未保存配置（手动加载或新增软件）" : "配置：" + Path.GetFileName(configurationPath);
            configurationPathLabel.AccessibleDescription = configurationPath ?? configurationPathLabel.Text;
            configurationTips.SetToolTip(configurationPathLabel, configurationPath ?? "首次保存可选择任意位置；每次启动手动加载。");
            Text = configurationPath == null ? "软件安装助理 — 新配置" : $"软件安装助理 — {Path.GetFileName(configurationPath)}";
        }

        private bool SaveConfiguration(bool saveAs)
        {
            string? path = configurationPath;
            if (saveAs || path == null)
            {
                using var dialog = new SaveFileDialog { Filter = "安装配置 (*.json)|*.json", DefaultExt = "json", AddExtension = true,
                    FileName = path == null ? "软件安装配置.json" : Path.GetFileName(path), Title = "选择配置文件保存位置" };
                if (path != null) dialog.InitialDirectory = Path.GetDirectoryName(path);
                if (dialog.ShowDialog(this) != DialogResult.OK) return false;
                path = dialog.FileName;
            }
            try
            {
                var configuration = CaptureConfiguration();
                ConfigurationStore.Save(path, configuration);
                configurationPath = path; savedConfiguration = ConfigurationStore.Serialize(configuration);
                UpdateConfigurationPath(); return true;
            }
            catch (Exception ex) { ErrorHandler.ShowError("保存配置失败：" + ex.Message); return false; }
        }

        private bool ConfirmConfigurationChange()
        {
            if (ConfigurationStore.Serialize(CaptureConfiguration()) == savedConfiguration) return true;
            var choice = MessageBox.Show(this, "当前配置有未保存的修改，是否保存？", "保存配置", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            return choice == DialogResult.No || (choice == DialogResult.Yes && SaveConfiguration(false));
        }

        private void LayoutBottomControls()
        {
            if (!_bottomControls.Any()) return;

            var rows = new Control[][] { [schemeComboBox, saveSchemeButton, deleteSchemeButton], [selectionSummary, manualInstallButton, autoInstallButton, buildBundleButton] };
            for (int row = 0; row < rows.Length; row++)
            {
                int currentX = 16;
                foreach (var control in rows[row])
                {
                    control.Location = new Point(currentX, 12 + row * 42);
                    currentX += control.Width + 12;
                }
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
            bool previousUpdating = isUpdatingByCode;
            isUpdatingByCode = true;
            try
            {
                categoryTreeView.Nodes.Clear();
                TreeNode allSoftwareNode = new TreeNode(AllSoftwareNodeText);
                categoryTreeView.Nodes.Add(allSoftwareNode);

                if (_softwareManager == null) return;

                var categoryNodes = CategoryNames.Normalize(_softwareManager.ActiveCategories.Concat(_softwareManager.ActiveSoftwareItems.Select(i => i.Category)))
                    .ToDictionary(name => name, name => new TreeNode(name), StringComparer.OrdinalIgnoreCase);

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

                foreach (TreeNode categoryNode in allSoftwareNode.Nodes)
                {
                    UpdateParentNodeCheckState(categoryNode);
                }
                UpdateParentNodeCheckState(allSoftwareNode);

                categoryTreeView.ExpandAll();
                UpdateSoftwareListView();
            }
            finally { isUpdatingByCode = previousUpdating; }
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
            bool allChildrenChecked = parentNode.Nodes.Count > 0 && parentNode.Nodes.Cast<TreeNode>().All(n => n.Checked);
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
            selectionSummary.Text = $"已添加 {_softwareManager?.ActiveSoftwareItems.Count ?? 0} 款 · 已勾选 {softwareListView.Items.Count} 款";
            LayoutBottomControls();
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
                using (var form = new SoftwareEditForm(itemToEdit, _softwareManager?.ActiveSoftwareItems, _softwareManager?.ActiveCategories))
                {
                    if (form.ShowDialog(this) == DialogResult.OK)
                    {
                        _softwareManager?.AddCategory(itemToEdit.Category);
                        RebuildTreeView();
                        UpdateSoftwareListView();
                    }
                }
            }
        }

        private void AddCategoryButton_Click(object? sender, EventArgs e)
        {
            if (_softwareManager == null) return;
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
                            try
                            {
                                _softwareManager.AddCategory(categoryName);
                                RebuildTreeView();
                            }
                            catch (Exception ex) { ErrorHandler.ShowError("保存分类失败：" + ex.Message); }
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
                    try { _softwareManager.RemoveCategory(selectedNode.Text); }
                    catch (Exception ex) { ErrorHandler.ShowError("删除分类失败：" + ex.Message); return; }
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
                openFileDialog.Filter = "安装文件 (*.exe;*.msi)|*.exe;*.msi";
                openFileDialog.Title = "请选择软件安装包";

                if (openFileDialog.ShowDialog(this) == DialogResult.OK)
                {
                    var newSoftware = new SoftwareItem
                    {
                        FilePath = openFileDialog.FileName,
                        Name = Path.GetFileNameWithoutExtension(openFileDialog.FileName),
                        Version = string.Empty,
                        Size = string.Empty,
                        Description = string.Empty,
                        Category = categoryName
                    };

                    try
                    {
                        var fileInfo = new FileInfo(newSoftware.FilePath);
                        var versionInfo = FileVersionInfo.GetVersionInfo(newSoftware.FilePath);
                        newSoftware.Version = (versionInfo.FileVersion ?? "N/A").Trim();
                        newSoftware.Size = $"{(fileInfo.Length / 1024.0 / 1024.0):F2} MB";
                        newSoftware.Description = (versionInfo.FileDescription ?? string.Empty).Trim();
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"无法读取文件信息: {ex.Message}");
                        newSoftware.Name = Path.GetFileNameWithoutExtension(newSoftware.FilePath);
                    }

                    using (var form = new SoftwareEditForm(newSoftware, _softwareManager.ActiveSoftwareItems, _softwareManager.ActiveCategories))
                    {
                        form.Text = "新增软件";
                        if (form.ShowDialog(this) == DialogResult.OK)
                        {
                            var finalSoftware = form.SoftwareItem;
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

            autoInstallButton.Enabled = manualInstallButton.Enabled = buildBundleButton.Enabled = false;
            try
            {
                var snapshot = _softwareManager.ActiveSoftwareItems.Select(s => s.Clone()).ToDictionary(s => s.Id);
                var descriptions = await Task.Run(() => snapshot.Values.Select(BundleBuilder.Describe).ToList());
                var items = InstallationEngine.IncludeDependencies(descriptions, selectedSoftware.Select(s => s.Id));
                var logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SoftwareInstaller", "Logs", Guid.NewGuid().ToString("N"));
                var reporter = new Progress<InstallUpdate>(p => Text = "软件安装助理 · " + p.Message);
                var results = await Task.Run(() => InstallationEngine.RunAsync(items, "", logDirectory, reporter,
                    resolveInstaller: s => snapshot[s.Id].FilePath!, requestElevation: true));
                MessageBox.Show(this, $"已处理 {results.Count}/{items.Count} 项，完成 {results.Count(r => r.State is InstallState.Succeeded or InstallState.Skipped or InstallState.RebootRequired)} 项。\n" +
                    string.Join("\n", results.Select(r => r.Name + "：" + r.Message)) + $"\n\n日志：{logDirectory}", "安装结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { ErrorHandler.ShowError(ex.Message); }
            finally { autoInstallButton.Enabled = manualInstallButton.Enabled = buildBundleButton.Enabled = true; Text = "软件安装助理"; }
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

                    try { _schemeHandler.SaveScheme(schemeName, _softwareManager.ActiveSoftwareItems, _softwareManager.ActiveCategories); }
                    catch (Exception ex) { ErrorHandler.ShowError("保存方案失败：" + ex.Message); return; }
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
                _softwareManager.ActiveCategories.Clear();
                foreach (var category in _schemeHandler.GetCategories(selectedScheme)) _softwareManager.AddCategory(category);
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
            if (_softwareManager != null && _schemeHandler != null && !ConfirmConfigurationChange()) e.Cancel = true;
        }
    }
}
