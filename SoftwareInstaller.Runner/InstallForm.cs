using System.Diagnostics;
using System.Security.Principal;
using SoftwareInstaller.Packaging;
using SoftwareInstaller.Models;
using SoftwareInstaller.Utils;

namespace SoftwareInstaller.Runner;

public sealed class InstallForm : Form
{
    private readonly ComboBox schemes = new() { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox categories = new() { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Margin = new(0) };
    private readonly CheckBox selectCategory = new() { AutoSize = true, AutoCheck = false, ThreeState = true, Anchor = AnchorStyles.Left, Enabled = false,
        Margin = new(0, 0, 8, 0), AccessibleName = "全选或取消选择当前范围内的软件" };
    private readonly ToolTip selectionTips = new();
    private readonly Dictionary<string, ListViewItem> rows = new();
    private readonly HashSet<string> selectedIds = new();
    private bool updatingSelection, initializing;
    private readonly ImageList softwareIcons = new() { ColorDepth = ColorDepth.Depth32Bit, ImageSize = new(20, 20) };
    private readonly ListView list = new() { Dock = DockStyle.Fill, CheckBoxes = true, View = View.Details, FullRowSelect = true, HideSelection = false, ShowItemToolTips = true };
    private readonly Label status = new() { Dock = DockStyle.Bottom, Height = 56, Padding = new(8), Text = "正在读取安装清单，请稍候…" };
    private readonly Label heading = new() { Dock = DockStyle.Top, Height = 80, Padding = new(12), Text = "软件安装包" };
    private readonly ProgressBar progress = new() { Dock = DockStyle.Bottom, Height = 16 };
    private readonly Button install = new() { Text = "一键安装", AutoSize = true, Enabled = false }, cancel = new() { Text = "取消安装", AutoSize = true, Enabled = false }, retry = new() { Text = "重试失败项", AutoSize = true, Enabled = false }, logs = new() { Text = "查看日志", AutoSize = true, Enabled = false };
    private BundleReader? package;
    private CancellationTokenSource? cancellation;
    private InstallationControl? installationControl;
    private readonly string[]? startupSelection;
    private readonly string packagePath;
    private readonly CancellationTokenSource loading = new();
    private readonly List<string> retained = new();
    private List<InstallResult> results = new();
    private bool running, uncertain, closeAfterRun;
    private string? logDirectory;
    internal bool PackageReady => package != null && install.Enabled;

    public InstallForm(string packagePath, string[]? startupSelection = null)
    {
        AppBranding.Apply(this);
        this.packagePath = Path.GetFullPath(packagePath);
        this.startupSelection = startupSelection;
        Text = "软件合集安装"; Size = new(920, 650); MinimumSize = new(760, 500); StartPosition = FormStartPosition.CenterScreen;
        list.SmallImageList = softwareIcons;
        _ = softwareIcons.Handle;
        list.Disposed += (_, _) => { softwareIcons.Dispose(); selectionTips.Dispose(); };
        list.Columns.Add("软件", 170); list.Columns.Add("版本", 70); list.Columns.Add("状态", 110); list.Columns.Add("说明", 270);
        list.Columns.Add("分类", 100); list.Columns.Add("默认安装位置", 190);
        list.Resize += (_, _) =>
        {
            int[] widths = [150, 60, 120, Math.Max(120, list.ClientSize.Width - 625), 80, 190];
            for (int i = 0; i < widths.Length; i++) list.Columns[i].Width = widths[i];
        };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Margin = new(0), WrapContents = false, FlowDirection = FlowDirection.RightToLeft };
        buttons.Controls.Add(install); buttons.Controls.Add(cancel); buttons.Controls.Add(retry); buttons.Controls.Add(logs);
        var footer = new TableLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, ColumnCount = 2, RowCount = 1, Padding = new(12) };
        footer.ColumnStyles.Add(new(SizeType.Percent, 100));
        footer.ColumnStyles.Add(new(SizeType.AutoSize));
        footer.Controls.Add(new Label { Text = AppBranding.AuthorCredit, ForeColor = SystemColors.GrayText, AutoEllipsis = true, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
        footer.Controls.Add(buttons, 1, 0);
        var header = new Panel { Dock = DockStyle.Top, Height = heading.Height };
        heading.Dock = DockStyle.Fill;
        header.Controls.Add(heading);
        header.Controls.Add(new BrandLogo { Dock = DockStyle.Left, Width = 80, Padding = new(8), AccessibleName = "软件安装助理标识" });
        var categoryBar = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Height = 36, ColumnCount = 2, RowCount = 1, Padding = new(8, 4, 8, 4) };
        categoryBar.RowStyles.Add(new(SizeType.AutoSize));
        categoryBar.ColumnStyles.Add(new(SizeType.AutoSize));
        categoryBar.ColumnStyles.Add(new(SizeType.Percent, 100));
        categoryBar.Controls.Add(selectCategory, 0, 0); categoryBar.Controls.Add(categories, 1, 0);
        Controls.Add(list); Controls.Add(categoryBar); Controls.Add(schemes); Controls.Add(header); Controls.Add(status); Controls.Add(progress); Controls.Add(footer);
        schemes.SelectedIndexChanged += (_, _) => { if (!initializing) ApplyScheme(); };
        categories.SelectedIndexChanged += (_, _) => { if (!initializing) RenderList(); };
        selectCategory.Click += (_, _) =>
        {
            if (package == null || running || uncertain) return;
            bool select = selectCategory.CheckState != CheckState.Checked;
            foreach (ListViewItem row in list.Items)
                if (select) selectedIds.Add((string)row.Tag!); else selectedIds.Remove((string)row.Tag!);
            SynchronizeChecks();
        };
        list.ItemChecked += (_, e) =>
        {
            if (updatingSelection) return;
            if (e.Item.Checked) selectedIds.Add((string)e.Item.Tag!); else selectedIds.Remove((string)e.Item.Tag!);
            UpdateCategorySelection();
        };
        install.Click += async (_, _) => await BeginInstallAsync();
        cancel.Click += (_, _) => CancelInstallation();
        retry.Click += async (_, _) =>
        {
            var ids = results.Where(r => r.State == InstallState.Failed).Select(r => r.Id).ToHashSet();
            SetSelection(ids);
            await BeginInstallAsync();
        };
        logs.Click += (_, _) => { if (logDirectory != null) Process.Start(new ProcessStartInfo("explorer.exe", $"\"{logDirectory}\"") { UseShellExecute = true }); };
        Shown += async (_, _) => await LoadPackageAsync();
        FormClosing += (_, e) =>
        {
            if (!running) return;
            e.Cancel = true;
            closeAfterRun = true;
            CancelInstallation();
        };
        FormClosed += (_, _) => { loading.Cancel(); package?.Dispose(); cancellation?.Dispose(); installationControl?.Dispose(); };
    }

    private void CancelInstallation()
    {
        if (!running) return;
        cancellation?.Cancel(); cancel.Enabled = false;
        status.Text = "正在取消当前安装并停止后续任务，请稍候…";
    }

    private async Task LoadPackageAsync()
    {
        progress.Style = ProgressBarStyle.Marquee;
        try
        {
            var loaded = await Task.Run(() => BundleReader.OpenAsync(packagePath, loading.Token));
            if (IsDisposed) { loaded.Dispose(); return; }
            package = loaded;
            var startupPlan = startupSelection == null ? null : package.Plan.Select(startupSelection);
            heading.Text = $"{package.Manifest.Name}\n包含 {package.Manifest.Software.Count} 款软件 · 制作时间 {package.Manifest.CreatedAt:yyyy-MM-dd HH:mm}";
            int online = package.Manifest.Software.Count(s => s.RequiresNetwork);
            if (online > 0) heading.Text += $" · {online} 款需要联网";
            int loadedRows = 0;
            foreach (var item in package.Plan.Items)
            {
                using var icon = SoftwareIcons.FromPng(item.IconPng);
                softwareIcons.Images.Add(icon);
                var row = new ListViewItem(item.Name) { Tag = item.Id, ToolTipText = item.Name, ImageIndex = softwareIcons.Images.Count - 1 };
                row.SubItems.Add(item.Version); row.SubItems.Add("待安装");
                row.SubItems.Add(string.Join("；", new[] { item.StartInteractively ? "将打开安装界面" : "", item.RequiresNetwork ? "需要联网" : "" }.Where(s => s.Length > 0))); row.SubItems.Add(item.Category);
                row.SubItems.Add(string.IsNullOrWhiteSpace(item.DefaultInstallDirectory) ? "跟随安装器默认" : item.DefaultInstallDirectory +
                    (item.DirectoryMode == InstallDirectoryMode.Automatic ? "（安装时选择）" : ""));
                rows.Add(item.Id, row);
                if (++loadedRows % 24 == 0)
                {
                    await Task.Delay(1, loading.Token);
                }
            }
            initializing = true;
            try
            {
                categories.Items.Add(new CategoryChoice(null));
                foreach (var name in CategoryNames.Normalize(package.Manifest.Categories.Concat(package.Manifest.Software.Select(s => s.Category))))
                    categories.Items.Add(new CategoryChoice(name));
                categories.SelectedIndex = 0;
                schemes.Items.AddRange(package.Manifest.Schemes.Keys.Cast<object>().ToArray());
                schemes.SelectedItem = package.Manifest.DefaultScheme;
                selectedIds.UnionWith(startupPlan != null ? startupPlan.Items.Select(s => s.Id) : package.Manifest.Schemes[package.Manifest.DefaultScheme]);
            }
            finally { initializing = false; }
            RenderList();
            status.Text = "请选择方案并点击一键安装。需要操作时会打开软件安装界面，请按提示完成。";
            install.Enabled = true;
            if (startupPlan != null)
            {
                if (!IsAdministrator()) throw new InvalidOperationException("管理员授权尚未完成。");
                await BeginInstallAsync(startupPlan);
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception ex) { if (!IsDisposed) { status.Text = "无法打开安装包：" + ex.Message; install.Enabled = false; } }
        finally { if (!IsDisposed) progress.Style = ProgressBarStyle.Blocks; }
    }

    private void ApplyScheme()
    {
        if (package == null || schemes.SelectedItem is not string name) return;
        SetSelection(package.Manifest.Schemes[name]);
    }

    private sealed record CategoryChoice(string? Name)
    {
        public override string ToString() => Name == null ? "全部分类" : "分类：" + Name;
    }

    private void SetSelection(IEnumerable<string> ids)
    {
        selectedIds.Clear(); selectedIds.UnionWith(ids);
        SynchronizeChecks();
    }

    private void SynchronizeChecks()
    {
        updatingSelection = true;
        list.BeginUpdate();
        try
        {
            foreach (ListViewItem row in list.Items)
            {
                bool check = selectedIds.Contains((string)row.Tag!);
                if (row.Checked != check) row.Checked = check;
            }
        }
        finally { list.EndUpdate(); updatingSelection = false; }
        UpdateCategorySelection();
    }

    private void RenderList()
    {
        var category = (categories.SelectedItem as CategoryChoice)?.Name;
        updatingSelection = true;
        list.BeginUpdate();
        try
        {
            list.Items.Clear();
            foreach (var pair in rows)
            {
                if (category != null && !pair.Value.SubItems[4].Text.Equals(category, StringComparison.OrdinalIgnoreCase)) continue;
                pair.Value.Checked = selectedIds.Contains(pair.Key);
                list.Items.Add(pair.Value);
            }
        }
        finally { list.EndUpdate(); updatingSelection = false; }
        UpdateCategorySelection();
    }

    private void UpdateCategorySelection()
    {
        int total = list.Items.Count;
        int selected = list.Items.Cast<ListViewItem>().Count(row => selectedIds.Contains((string)row.Tag!));
        selectCategory.CheckState = selected == 0 ? CheckState.Unchecked : selected == total ? CheckState.Checked : CheckState.Indeterminate;
        selectCategory.Enabled = package != null && !running && !uncertain && total > 0;
        selectionTips.SetToolTip(selectCategory, selected == total && total > 0 ? "取消选择当前范围内的全部软件" : "选中当前范围内的全部软件");
    }

    private static bool IsAdministrator() => new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);

    private async Task BeginInstallAsync(InstallationPlan? preparedPlan = null)
    {
        if (package == null || running || uncertain) return;
        var selected = selectedIds.ToList();
        if (selected.Count == 0) { status.Text = "请至少勾选一款软件。"; return; }
        var plan = preparedPlan ?? package.Plan.Select(selected);
        var items = plan.Items;
        var itemIds = items.Select(s => s.Id).ToHashSet();
        selectedIds.UnionWith(itemIds);
        SynchronizeChecks();
        if (!IsAdministrator())
        {
            try
            {
                Process.Start(new ProcessStartInfo(packagePath, "--install " + string.Join(",", items.Select(s => s.Id))) { UseShellExecute = true, Verb = "runas" });
                Close();
            }
            catch (Exception ex) { status.Text = "未能取得管理员权限：" + ex.Message; }
            return;
        }
        running = true; install.Enabled = retry.Enabled = schemes.Enabled = categories.Enabled = selectCategory.Enabled = list.Enabled = false; cancel.Enabled = true;
        results = new(); progress.Value = 0;
        cancellation = new();
        installationControl = new();
        var work = Path.Combine(Path.GetTempPath(), "SIA-install-" + Guid.NewGuid().ToString("N"));
        logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SoftwareInstaller", "Logs", DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8));
        try
        {
            Directory.CreateDirectory(logDirectory); logs.Enabled = true;
            foreach (ListViewItem row in rows.Values)
                if (itemIds.Contains((string)row.Tag!)) { row.Checked = true; row.SubItems[2].Text = "等待执行"; row.SubItems[3].Text = ""; }
            var extraction = new Progress<BundleProgress>(p =>
            {
                status.Text = $"正在解包当前软件 {p.Completed}/{p.Total} · {p.Message}";
                progress.Style = ProgressBarStyle.Blocks;
                progress.Value = p.TotalBytes > 0 ? (int)Math.Max(0, Math.Min(100, p.BytesCompleted * 100 / p.TotalBytes)) : p.Total == 0 ? 0 : p.Completed * 100 / p.Total;
            });
            var reporter = new Progress<InstallUpdate>(p =>
            {
                status.Text = $"已处理 {p.Completed}/{p.Total} 款 · {p.StateText} · {p.Message}";
                progress.Style = p.Result == null && p.Phase != InstallPhase.Preparing ? ProgressBarStyle.Marquee : ProgressBarStyle.Blocks;
                if (p.Phase == InstallPhase.Preparing) progress.Value = 0;
                if (p.Result != null) progress.Value = p.Total == 0 ? 0 : p.Completed * 100 / p.Total;
                var row = rows[p.Id];
                row.SubItems[2].Text = p.StateText;
                row.SubItems[3].Text = p.Message;
            });
            results = await Task.Run(() => InstallationEngine.RunAsync(plan, work, logDirectory, reporter, cancellation.Token, control: installationControl,
                prepareInstaller: (item, token) => package.ExtractAsync([item], work, extraction, token), cancelCurrent: true));
            uncertain = results.Any(r => r.State == InstallState.Uncertain || r.ExitCode == 1641);
            var finished = results.Select(r => r.Id).ToHashSet();
            foreach (ListViewItem row in rows.Values)
                if (itemIds.Contains((string)row.Tag!) && !finished.Contains((string)row.Tag!)) { row.SubItems[2].Text = "未执行"; row.SubItems[3].Text = "队列已停止"; }
            status.Text = $"已处理 {results.Count}/{items.Count} 项：完成 {results.Count(r => r.State == InstallState.Succeeded)}，跳过 {results.Count(r => r.State == InstallState.Skipped)}，失败 {results.Count(r => r.State == InstallState.Failed)}，取消 {results.Count(r => r.State == InstallState.Cancelled)}，待检查 {results.Count(r => r.State == InstallState.Uncertain)}，需重启 {results.Count(r => r.State == InstallState.RebootRequired)}，未执行 {items.Count - results.Count}。";
            if (results.Any(r => r.State == InstallState.RebootRequired)) status.Text += " 部分软件需要重启。";
            if (uncertain) { retained.Add(work); status.Text += " 队列已暂停；请检查仍在运行的安装器，资源暂时保留。"; }
        }
        catch (OperationCanceledException) { status.Text = "已取消准备安装。"; }
        catch (Exception ex)
        {
            // Preserve files if an unexpected failure occurs after an installer may have started.
            retained.Add(work); uncertain = true;
            status.Text = "安装流程已停止：" + ex.Message;
            try { await RunnerRuntime.WriteTextAsync(Path.Combine(logDirectory, "error.log"), ex.ToString()); }
            catch (Exception logError) when (logError is IOException or UnauthorizedAccessException) { status.Text += " 日志无法写入。"; }
        }
        finally
        {
            progress.Style = ProgressBarStyle.Blocks;
            progress.Value = items.Count == 0 ? 0 : results.Count * 100 / items.Count;
            if (!retained.Contains(work) && Directory.Exists(work))
            {
                try { await Task.Run(() => Directory.Delete(work, true)); }
                catch (IOException) { retained.Add(work); }
                catch (UnauthorizedAccessException) { retained.Add(work); }
            }
            if (retained.Count > 0)
            {
                try { await RunnerRuntime.WriteTextAsync(Path.Combine(logDirectory, "retained-resources.txt"), string.Join(Environment.NewLine, retained)); }
                catch (Exception logError) when (logError is IOException or UnauthorizedAccessException) { status.Text += " 资源目录：" + work; }
                status.Text += " 保留目录见日志。";
            }
            running = false; cancellation.Dispose(); cancellation = null; cancel.Enabled = false;
            installationControl.Dispose(); installationControl = null;
            install.Enabled = schemes.Enabled = categories.Enabled = list.Enabled = !uncertain;
            UpdateCategorySelection();
            retry.Enabled = !uncertain && results.Any(r => r.State == InstallState.Failed);
            if (closeAfterRun) BeginInvoke(new Action(Close));
        }
    }

}
