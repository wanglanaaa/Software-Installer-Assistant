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
        private readonly Button buildBundleButton = new() { Text = "生成单文件安装包" };
        private readonly Button projectButton = new() { Text = "更新版本" };
        private readonly Label selectionSummary = new() { AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft };
        private readonly BrandLogo cornerLogo = new() { AccessibleName = "软件安装助理标识" };
        private const int ConfigurationBarHeight = 64;
        private const int PageMargin = 20, SidebarWidth = 200, ColumnGap = 12, ButtonHeight = 32, ButtonWidth = 96, ButtonGap = 8;
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
            InitializeSoftwareList();
            InitializeOrdering();
            InitializeSchemeManagement();
            RebuildTreeView();
        }

        private void InitializeCustomComponents()
        {
            installPathTextBox = new TextBox { Size = new System.Drawing.Size(150, 23), Text = DefaultInstallPath };
            browseButton = new Button { Text = "安装路径", Size = new System.Drawing.Size(80, 25) };
            manualInstallButton = new Button { Text = "手动安装" };
            autoInstallButton = new Button { Text = "自动安装" };
            schemeComboBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
            saveSchemeButton = new Button { Text = "保存方案" };
            deleteSchemeButton = new Button { Text = "删除方案" };
            addCategoryButton.Text = "新增分类"; deleteCategoryButton.Text = "删除分类";
            addSoftwareButton.Text = "新增软件"; deleteSoftwareButton.Text = "删除软件";

            foreach (var button in new[] { addCategoryButton, deleteCategoryButton, addSoftwareButton, deleteSoftwareButton,
                saveSchemeButton, deleteSchemeButton, manualInstallButton, autoInstallButton, projectButton }) StyleButton(button);
            StyleButton(buildBundleButton, true);
            projectButton.Click += (_, _) =>
            {
                try { Process.Start(new ProcessStartInfo("https://github.com/wanglanaaa/Software-Installer-Assistant") { UseShellExecute = true }); }
                catch (Exception ex) { ErrorHandler.ShowError("无法打开项目页面：" + ex.Message); }
            };
            configurationTips.SetToolTip(projectButton, "打开项目源码、版本下载和问题反馈页面");
            buildBundleButton.Click += (_, _) =>
            {
                if (_softwareManager == null || _schemeHandler == null) return;
                using var dialog = new BundleBuildForm(_softwareManager.ActiveSoftwareItems, _schemeHandler, _softwareManager.ActiveCategories);
                dialog.ShowDialog(this);
            };
            bottomPanel.Height = 108;
            bottomPanel.Dock = DockStyle.Bottom;
            MinimumSize = new Size(820, 650);
            // A global path is not supported by arbitrary third-party installers.
            installPathTextBox.Visible = browseButton.Visible = false;

            configurationTips.SetToolTip(cornerLogo, "软件安装助理");
            _bottomControls.AddRange(new Control[] { schemeComboBox, saveSchemeButton, deleteSchemeButton, projectButton, selectionSummary, manualInstallButton, autoInstallButton, buildBundleButton, cornerLogo });

            this.Load += OnMainFormLoad;
            this.Resize += (_, _) => LayoutContent();
            bottomPanel.Paint += BottomPanel_Paint;
            bottomPanel.Resize += OnBottomPanelResize;

            categoryTreeView.AfterCheck += CategoryTreeView_AfterCheck;
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
            int top = ConfigurationBarHeight + ButtonGap;
            int bottom = ClientSize.Height - bottomPanel.Height - 16;
            int listLeft = PageMargin + SidebarWidth + ColumnGap;
            int listWidth = Math.Max(300, ClientSize.Width - listLeft - PageMargin);
            softwareSearch.SetBounds(listLeft, top, Math.Max(120, listWidth - 168), 26);
            softwareCategory.SetBounds(listLeft + listWidth - 156, top, 156, 26);
            softwareListView.SetBounds(listLeft, top + 34, listWidth, Math.Max(120, bottom - top - 34 - ButtonHeight - ButtonGap));
            softwareMoveUp.SetBounds(listLeft, bottom - ButtonHeight, ButtonWidth, ButtonHeight);
            softwareMoveDown.SetBounds(listLeft + ButtonWidth + ButtonGap, bottom - ButtonHeight, ButtonWidth, ButtonHeight);
            if (softwareListView.Columns.Count == 6)
            {
                int[] widths = [48, 144, 64, 72, Math.Max(110, listWidth - 440), 88];
                for (int i = 0; i < widths.Length; i++) softwareListView.Columns[i].Width = widths[i];
            }
            int firstRow = bottom - ButtonHeight * 3 - ButtonGap * 2;
            categoryTreeView.SetBounds(PageMargin, top, SidebarWidth, Math.Max(100, firstRow - ColumnGap - top));
            treeMoveUp.SetBounds(PageMargin, firstRow, ButtonWidth, ButtonHeight);
            treeMoveDown.SetBounds(PageMargin + ButtonWidth + ButtonGap, firstRow, ButtonWidth, ButtonHeight);
            firstRow += ButtonHeight + ButtonGap;
            addCategoryButton.SetBounds(PageMargin, firstRow, ButtonWidth, ButtonHeight);
            deleteCategoryButton.SetBounds(PageMargin + ButtonWidth + ButtonGap, firstRow, ButtonWidth, ButtonHeight);
            addSoftwareButton.SetBounds(PageMargin, firstRow + ButtonHeight + ButtonGap, ButtonWidth, ButtonHeight);
            deleteSoftwareButton.SetBounds(PageMargin + ButtonWidth + ButtonGap, firstRow + ButtonHeight + ButtonGap, ButtonWidth, ButtonHeight);
        }

        private static void StyleButton(Button button, bool primary = false)
        {
            button.AutoSize = false;
            button.Size = new Size(ButtonWidth, ButtonHeight);
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.BorderColor = primary ? Color.FromArgb(37, 99, 235) : Color.FromArgb(203, 213, 225);
            button.BackColor = primary ? Color.FromArgb(37, 99, 235) : Color.FromArgb(248, 250, 252);
            button.ForeColor = primary ? Color.White : Color.FromArgb(51, 65, 85);
            button.FlatAppearance.MouseOverBackColor = primary ? Color.FromArgb(29, 78, 216) : Color.FromArgb(226, 232, 240);
            button.FlatAppearance.MouseDownBackColor = primary ? Color.FromArgb(30, 64, 175) : Color.FromArgb(203, 213, 225);
            button.UseVisualStyleBackColor = false;
        }

        private void InitializeConfigurationBar()
        {
            var bar = new TableLayoutPanel { Dock = DockStyle.Top, Height = ConfigurationBarHeight, ColumnCount = 6, RowCount = 1, Padding = new(PageMargin, PageMargin, PageMargin, 12) };
            bar.RowStyles.Add(new(SizeType.Percent, 100));
            for (int i = 0; i < 5; i++) bar.ColumnStyles.Add(new(SizeType.Absolute, ButtonWidth + (i == 1 ? ColumnGap : ButtonGap)));
            bar.ColumnStyles.Add(new(SizeType.Percent, 100));
            void AddButton(string text, int column, Action action)
            {
                var button = new Button { Text = text, Dock = DockStyle.Fill, Margin = new(0, 0, column == 1 ? ColumnGap : ButtonGap, 0) };
                StyleButton(button);
                button.Click += (_, _) => action(); bar.Controls.Add(button, column, 0);
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
            AddButton("修复路径…", 4, RepairResourcePaths);
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
            configurationRevision++;
            softwareSearch.Clear(); softwareCategory.SelectedIndex = -1;
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

            const int firstRow = 12, secondRow = 56;
            int actionsLeft = PageMargin + SidebarWidth + ColumnGap;
            schemeComboBox.Width = SidebarWidth;
            schemeComboBox.Location = new Point(PageMargin, firstRow + (ButtonHeight - schemeComboBox.Height) / 2);
            saveSchemeButton.SetBounds(actionsLeft, firstRow, ButtonWidth, ButtonHeight);
            deleteSchemeButton.SetBounds(actionsLeft + ButtonWidth + ButtonGap, firstRow, ButtonWidth, ButtonHeight);
            projectButton.SetBounds(actionsLeft + (ButtonWidth + ButtonGap) * 2, firstRow, 160, ButtonHeight);
            selectionSummary.SetBounds(PageMargin, secondRow, SidebarWidth, ButtonHeight);
            manualInstallButton.SetBounds(actionsLeft, secondRow, ButtonWidth, ButtonHeight);
            autoInstallButton.SetBounds(actionsLeft + ButtonWidth + ButtonGap, secondRow, ButtonWidth, ButtonHeight);
            buildBundleButton.SetBounds(actionsLeft + (ButtonWidth + ButtonGap) * 2, secondRow, 160, ButtonHeight);
            cornerLogo.SetBounds(bottomPanel.ClientSize.Width - PageMargin - 80, bottomPanel.ClientSize.Height - PageMargin - 80, 80, 80);
        }

        private void BottomPanel_Paint(object? sender, PaintEventArgs e)
        {
            using var pen = new Pen(Color.Gainsboro, 1);
            e.Graphics.DrawLine(pen, PageMargin, 0, bottomPanel.Width - PageMargin, 0);
        }

        private void InitializeSchemeManagement()
        {
            PopulateSchemeComboBox();
        }

        private void RebuildTreeView()
        {
            var selectedNode = categoryTreeView.SelectedNode;
            string? selectedSoftwareId = (selectedNode?.Tag as SoftwareItem)?.Id;
            string? selectedCategory = selectedNode?.Parent != null && selectedNode.Tag is not SoftwareItem ? selectedNode.Text : null;
            bool selectedRoot = selectedNode?.Parent == null && selectedNode != null;
            bool previousUpdating = isUpdatingByCode;
            isUpdatingByCode = true;
            try
            {
                categoryTreeView.Nodes.Clear();
                TreeNode allSoftwareNode = new TreeNode(AllSoftwareNodeText);
                categoryTreeView.Nodes.Add(allSoftwareNode);

                if (_softwareManager == null) return;
                RefreshSoftwareCategories();

                var categoryNodes = CategoryNames.Normalize(_softwareManager.ActiveCategories.Concat(_softwareManager.ActiveSoftwareItems.Select(i => i.Category)))
                    .ToDictionary(name => name, name => new TreeNode(name), StringComparer.OrdinalIgnoreCase);

                foreach (var node in categoryNodes.Values)
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
                categoryTreeView.SelectedNode = selectedRoot ? allSoftwareNode : allSoftwareNode.Nodes.Cast<TreeNode>()
                    .SelectMany(n => new[] { n }.Concat(n.Nodes.Cast<TreeNode>()))
                    .FirstOrDefault(n => selectedSoftwareId != null ? (n.Tag as SoftwareItem)?.Id == selectedSoftwareId
                        : selectedCategory != null && n.Tag is not SoftwareItem && n.Text == selectedCategory);
                UpdateSoftwareListView();
                UpdateOrderingButtons();
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
            if (updatingSoftwareList || _softwareManager == null) return;
            updatingSoftwareList = true; softwareListView.BeginUpdate();
            try
            {
                var selectedIds = softwareListView.SelectedItems.Cast<ListViewItem>().Select(r => ((SoftwareItem)r.Tag!).Id).ToHashSet();
                string? focusedId = (softwareListView.FocusedItem?.Tag as SoftwareItem)?.Id;
                softwareListView.Items.Clear(); int index = 1;
                string query = softwareSearch.Text.Trim(); string? category = (softwareCategory.SelectedItem as CategoryFilter)?.Name;
                foreach (var item in _softwareManager.ActiveSoftwareItems)
                {
                    if (category != null && !item.Category.Equals(category, StringComparison.OrdinalIgnoreCase)) continue;
                    if (query.Length > 0 && !new[] { item.Name, item.Version, item.Description, item.Category }
                        .Any(value => value?.Contains(query, StringComparison.OrdinalIgnoreCase) == true)) continue;
                    var row = new ListViewItem((index++).ToString()) { Tag = item, Checked = item.IsSelected };
                    row.SubItems.AddRange([item.Name, item.Version, item.Size, item.Description, item.Category]);
                    softwareListView.Items.Add(row);
                    row.Selected = selectedIds.Contains(item.Id);
                    row.Focused = item.Id == focusedId;
                }
                softwareListView.FocusedItem?.EnsureVisible();
                UpdateSelectionSummary();
                UpdateOrderingButtons();
            }
            finally { softwareListView.EndUpdate(); updatingSoftwareList = false; }
        }

        private void CategoryTreeView_NodeMouseDoubleClick(object? sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node?.Tag is SoftwareItem itemToEdit) EditSoftware(itemToEdit);
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

        private async void AddSoftwareButton_Click(object? sender, EventArgs e)
        {
            using var dialog = new OpenFileDialog { Filter = "安装文件 (*.exe;*.msi)|*.exe;*.msi", Title = "选择一个或多个软件安装包", Multiselect = true };
            if (dialog.ShowDialog(this) == DialogResult.OK) await ImportSoftwareAsync(dialog.FileNames);
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
                var reporter = new Progress<InstallUpdate>(p => Text = $"软件安装助理 · 已处理 {p.Completed}/{p.Total} · {p.StateText} · {p.Message}");
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
