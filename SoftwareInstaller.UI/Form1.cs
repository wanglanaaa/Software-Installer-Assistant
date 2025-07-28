using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using SoftwareInstaller.Models;

namespace SoftwareInstaller.UI
{
    public partial class Form1 : Form
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

        private List<Control> bottomControls = new List<Control>(); // Declare as class member

        private List<SoftwareItem> softwareItems = new List<SoftwareItem>();
        private Dictionary<string, List<SoftwareItem>> savedSchemes = new Dictionary<string, List<SoftwareItem>>();
        private string schemeFilePath = string.Empty;
        private bool isUpdatingByCode = false;

        public Form1()
        {
            InitializeComponent();
            InitializeCustomComponents();
            InitializeSchemeManagement();
            PopulateDummyData();
        }

        private void InitializeCustomComponents()
        {
            // =================================================================
            // UI Modernization (Edge Style)
            // =================================================================

            this.BackColor = Color.White;
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(0)));
            this.Text = "软件安装助理";
            this.Size = new System.Drawing.Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;

            // --- Top Controls ---
            categoryTreeView = new TreeView
            {
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(200, 380),
                BorderStyle = BorderStyle.FixedSingle,
                CheckBoxes = true,
                ShowLines = false // Cleaner look
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

            // --- Bottom Action Bar ---
            Panel bottomPanel = new Panel
            {
                Size = new System.Drawing.Size(this.ClientSize.Width, 80),
                BackColor = Color.White,
                Dock = DockStyle.Bottom
            };
            bottomPanel.Paint += (s, e) => e.Graphics.DrawLine(new Pen(Color.Gainsboro, 1), 0, 0, bottomPanel.Width, 0);

            // --- Controls inside Bottom Panel ---
            // Initialize controls with fixed sizes
            installPathTextBox = new TextBox { Size = new System.Drawing.Size(150, 23), Text = @"D:\Program Files" };
            browseButton = new Button { Text = "安装路径", Size = new System.Drawing.Size(80, 25) }; // Fixed size

            manualInstallButton = new Button { Text = "▶️ 手动安装", Size = new System.Drawing.Size(100, 25) }; // Fixed size
            autoInstallButton = new Button { Text = "⚡ 自动安装", Size = new System.Drawing.Size(100, 25) }; // Fixed size

            schemeComboBox = new ComboBox { Size = new System.Drawing.Size(120, 23), DropDownStyle = ComboBoxStyle.DropDownList };
            saveSchemeButton = new Button { Text = "保存方案", Size = new System.Drawing.Size(80, 25) }; // Fixed size
            deleteSchemeButton = new Button { Text = "删除方案", Size = new System.Drawing.Size(80, 25) }; // Fixed size

            // Add to bottomControls list in desired order
            bottomControls.Add(installPathTextBox);
            bottomControls.Add(browseButton);
            bottomControls.Add(manualInstallButton);
            bottomControls.Add(autoInstallButton);
            bottomControls.Add(schemeComboBox);
            bottomControls.Add(saveSchemeButton);
            bottomControls.Add(deleteSchemeButton);

            // --- Styling & Layout Logic ---
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

            // Dynamic layout for bottom controls
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

            // Add all controls to the form
            this.Controls.Add(categoryTreeView);
            this.Controls.Add(softwareListView);
            this.Controls.Add(bottomPanel);
            utilityButtons.ForEach(btn => this.Controls.Add(btn));

            // --- Event Handlers ---
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
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string appFolderPath = Path.Combine(appDataPath, "SoftwareInstaller");
            Directory.CreateDirectory(appFolderPath);
            schemeFilePath = Path.Combine(appFolderPath, "schemes.json");
            LoadSchemesFromFile();
        }

        private void PopulateDummyData()
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

            AddCategory("办公软件");
            AddCategory("文字输入");
            AddCategory("维护工具");
            AddCategory("网络工具");

            softwareItems.AddRange(new List<SoftwareItem>
            {
                new SoftwareItem { Name = "WPS Office", Version = "11.1.0.9021", Size = "70MB", Description = "WPS Office是一款办公软件", IsSelected = false, Category = "办公软件" },
                new SoftwareItem { Name = "Office2016免混", Version = "", Size = "", Description = "", IsSelected = false, Category = "办公软件" },
                new SoftwareItem { Name = "System爱好者维护工具...", Version = "25.4.17.0", Size = "175MB", Description = "System爱好者社区打造的装", IsSelected = false, Category = "维护工具" },
                new SoftwareItem { Name = "搜狗拼音输入法", Version = "", Size = "", Description = "", IsSelected = false, Category = "文字输入" },
                new SoftwareItem { Name = "搜狗五笔输入法", Version = "", Size = "", Description = "", IsSelected = false, Category = "文字输入" },
                new SoftwareItem { Name = "微信正式版", Version = "4.0.3.40", Size = "210MB", Description = "跨平台通讯工具", IsSelected = false, Category = "网络工具" },
                new SoftwareItem { Name = "QQ9.7", Version = "", Size = "", Description = "", IsSelected = false, Category = "网络工具" },
                new SoftwareItem { Name = "爱奇艺", Version = "13.1.0.8958", Size = "64MB", Description = "爱奇艺,中国高品质视频娱", IsSelected = false, Category = "网络工具" },
                new SoftwareItem { Name = "360安全浏览器", Version = "15.2.6440.0", Size = "141MB", Description = "全面守护上网安全,防病毒网", IsSelected = false, Category = "维护工具" },
                new SoftwareItem { Name = "360压缩", Version = "4.0.0.1520", Size = "15MB", Description = "360压缩是新一代的压缩软", IsSelected = false, Category = "维护工具" },
                new SoftwareItem { Name = "360安全卫士", Version = "13.0.0.2258", Size = "94MB", Description = "安全卫士13.0全面安全、全", IsSelected = false, Category = "维护工具" }
            });

            foreach (var item in softwareItems)
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
            if (e.Node == null) return; // Fix CS8602: Ensure e.Node is not null

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
            string? categoryName = Interaction.InputBox("请输入新的分类名称:", "增加分类", "");
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

        private void DeleteCategoryButton_Click(object? sender, EventArgs e)
        {
            TreeNode? selectedNode = categoryTreeView.SelectedNode;
            if (selectedNode != null && selectedNode.Parent != null && selectedNode.Parent.Text == "所有软件")
            {
                if (MessageBox.Show($"确定要删除分类 '{selectedNode.Text}' 吗?\n这将同时删除该分类下的所有软件。", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    softwareItems.RemoveAll(item => item.Category == selectedNode.Text);
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

            // Prevent adding software directly under the "所有软件" root node
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
                    string? description = Interaction.InputBox("请编辑软件说明:", "增加软件", versionInfo.Comments ?? "");
                    string? silentArgs = Interaction.InputBox("请输入该软件的静默安装参数(可选):", "增加软件", "");

                    var newSoftware = new SoftwareItem
                    {
                        Name = name, Version = version, Size = size, Description = description ?? "",
                        Category = selectedNode.Text, FilePath = filePath, SilentInstallArgs = silentArgs, IsSelected = true
                    };
                    softwareItems.Add(newSoftware);

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
                    softwareItems.Remove(itemToRemove);
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
            ExecuteInstallation(false);
        }

        private void AutoInstallButton_Click(object? sender, EventArgs e)
        {
            ExecuteInstallation(true);
        }

        private void ExecuteInstallation(bool isAuto)
        {
            List<SoftwareItem> selectedSoftware = GetSelectedSoftwareItems();
            if (selectedSoftware.Count == 0)
            {
                MessageBox.Show("没有选择任何软件进行安装。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (var item in selectedSoftware)
            {
                if (string.IsNullOrEmpty(item.FilePath))
                {
                    MessageBox.Show($"软件 '{item.Name}' 未找到安装文件路径。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    continue;
                }

                try
                {
                    ProcessStartInfo startInfo = new ProcessStartInfo();
                    startInfo.FileName = item.FilePath;
                    if (isAuto)
                    {
                        if (string.IsNullOrEmpty(item.SilentInstallArgs))
                        {
                            MessageBox.Show($"软件 '{item.Name}' 未提供静默安装参数，无法自动安装。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            continue;
                        }
                        startInfo.Arguments = item.SilentInstallArgs;
                    }
                    Process.Start(startInfo);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"无法启动安装程序 '{item.Name}'.\n错误: {ex.Message}", "安装失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private List<SoftwareItem> GetSelectedSoftwareItems()
        {
            return softwareItems.Where(item => item.IsSelected).ToList();
        }

        private void SaveSchemeButton_Click(object? sender, EventArgs e)
        {
            string? schemeName = Interaction.InputBox("请输入方案名称:", "保存方案", "");
            if (string.IsNullOrWhiteSpace(schemeName)) return;

            if (savedSchemes.ContainsKey(schemeName))
            {
                if (MessageBox.Show("该方案名称已存在，要覆盖吗?", "确认覆盖", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                {
                    return;
                }
            }

            SyncModelFromTree(categoryTreeView.Nodes);
            savedSchemes[schemeName] = new List<SoftwareItem>(softwareItems.Select(item => item.Clone()));
            SaveSchemesToFile();
            UpdateSchemeComboBox();
            schemeComboBox.SelectedItem = schemeName;
        }

        private void DeleteSchemeButton_Click(object? sender, EventArgs e)
        {
            if (schemeComboBox.SelectedItem is string selectedScheme && !string.IsNullOrEmpty(selectedScheme))
            {
                if (MessageBox.Show($"确定要删除方案 '{selectedScheme}' 吗?", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    savedSchemes.Remove(selectedScheme);
                    SaveSchemesToFile();
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
                foreach (var item in softwareItems)
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

        private void LoadSchemesFromFile()
        {
            try
            {
                if (File.Exists(schemeFilePath))
                {
                    string jsonString = File.ReadAllText(schemeFilePath);
                    savedSchemes = JsonSerializer.Deserialize<Dictionary<string, List<SoftwareItem>>>(jsonString) ?? new Dictionary<string, List<SoftwareItem>>();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载方案文件失败.\n错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            UpdateSchemeComboBox();
        }

        private void SaveSchemesToFile()
        {
            try
            {
                string jsonString = JsonSerializer.Serialize(savedSchemes, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(schemeFilePath, jsonString);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存方案文件失败.\n错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            var categoryNodes = softwareItems.Select(i => i.Category).Distinct()
                .ToDictionary(name => name, name => new TreeNode(name));

            foreach (var node in categoryNodes.Values.OrderBy(n => n.Text))
            {
                allSoftwareNode.Nodes.Add(node);
            }

            foreach (var item in softwareItems)
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