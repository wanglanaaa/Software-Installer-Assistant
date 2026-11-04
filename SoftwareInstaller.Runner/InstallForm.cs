using System.Diagnostics;
using System.Security.Principal;
using SoftwareInstaller.Packaging;
using SoftwareInstaller.Models;

namespace SoftwareInstaller.Runner;

public sealed class InstallForm : Form
{
    private readonly ComboBox schemes = new() { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly ComboBox categories = new() { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly Dictionary<string, ListViewItem> rows = new();
    private readonly HashSet<string> selectedIds = new();
    private bool updatingSelection;
    private readonly ListView list = new() { Dock = DockStyle.Fill, CheckBoxes = true, View = View.Details, FullRowSelect = true, HideSelection = false, ShowItemToolTips = true };
    private readonly Label status = new() { Dock = DockStyle.Bottom, Height = 56, Padding = new(8), Text = "正在校验安装包，请稍候…" };
    private readonly Label heading = new() { Dock = DockStyle.Top, Height = 80, Padding = new(12), Text = "软件安装包" };
    private readonly ProgressBar progress = new() { Dock = DockStyle.Bottom, Height = 16 };
    private readonly Button install = new() { Text = "一键安装", AutoSize = true, Enabled = false }, cancel = new() { Text = "停止后续任务", AutoSize = true, Enabled = false }, retry = new() { Text = "重试失败项", AutoSize = true, Enabled = false }, logs = new() { Text = "查看日志", AutoSize = true, Enabled = false };
    private BundleReader? package;
    private CancellationTokenSource? cancellation;
    private readonly string[]? startupSelection;
    private readonly string packagePath;
    private readonly CancellationTokenSource loading = new();
    private readonly List<string> retained = new();
    private List<InstallResult> results = new();
    private bool running, uncertain;
    private string? logDirectory;
    internal bool PackageReady => package != null && install.Enabled;

    public InstallForm(string packagePath, string[]? startupSelection = null)
    {
        AppBranding.Apply(this);
        this.packagePath = Path.GetFullPath(packagePath);
        this.startupSelection = startupSelection;
        Text = "软件合集安装"; Size = new(920, 650); MinimumSize = new(760, 500); StartPosition = FormStartPosition.CenterScreen;
        list.Columns.Add("软件", 170); list.Columns.Add("版本", 70); list.Columns.Add("状态", 110); list.Columns.Add("说明", 270);
        list.Columns.Add("分类", 100); list.Columns.Add("默认安装位置", 190);
        list.Resize += (_, _) =>
        {
            int[] widths = [150, 60, 120, Math.Max(120, list.ClientSize.Width - 625), 80, 190];
            for (int i = 0; i < widths.Length; i++) list.Columns[i].Width = widths[i];
        };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, Padding = new(12), FlowDirection = FlowDirection.RightToLeft };
        buttons.Controls.Add(install); buttons.Controls.Add(cancel); buttons.Controls.Add(retry); buttons.Controls.Add(logs);
        var header = new Panel { Dock = DockStyle.Top, Height = heading.Height };
        heading.Dock = DockStyle.Fill;
        header.Controls.Add(heading);
        header.Controls.Add(new BrandLogo { Dock = DockStyle.Left, Width = 80, Padding = new(8), AccessibleName = "软件安装助理标识" });
        Controls.Add(list); Controls.Add(categories); Controls.Add(schemes); Controls.Add(header); Controls.Add(status); Controls.Add(progress); Controls.Add(buttons);
        schemes.SelectedIndexChanged += (_, _) => ApplyScheme();
        categories.SelectedIndexChanged += (_, _) => RenderList();
        list.ItemChecked += (_, e) =>
        {
            if (updatingSelection) return;
            if (e.Item.Checked) selectedIds.Add((string)e.Item.Tag!); else selectedIds.Remove((string)e.Item.Tag!);
        };
        install.Click += async (_, _) => await BeginInstallAsync();
        cancel.Click += (_, _) => { cancellation?.Cancel(); cancel.Enabled = false; status.Text = "将停止后续任务；当前安装会等待结束。"; };
        retry.Click += async (_, _) =>
        {
            var ids = results.Where(r => r.State == InstallState.Failed).Select(r => r.Id).ToHashSet();
            SetSelection(ids);
            await BeginInstallAsync();
        };
        logs.Click += (_, _) => { if (logDirectory != null) Process.Start(new ProcessStartInfo("explorer.exe", $"\"{logDirectory}\"") { UseShellExecute = true }); };
        Shown += async (_, _) => await LoadPackageAsync();
        FormClosing += (_, e) => { if (running) { e.Cancel = true; cancellation?.Cancel(); status.Text = "正在停止后续任务，请等待当前操作结束。"; } };
        FormClosed += (_, _) => { loading.Cancel(); package?.Dispose(); cancellation?.Dispose(); };
    }

    private async Task LoadPackageAsync()
    {
        progress.Style = ProgressBarStyle.Marquee;
        try
        {
            var loaded = await Task.Run(() => BundleReader.OpenAsync(packagePath, loading.Token));
            if (IsDisposed) { loaded.Dispose(); return; }
            package = loaded;
            heading.Text = $"{package.Manifest.Name}\n包含 {package.Manifest.Software.Count} 款软件 · 制作时间 {package.Manifest.CreatedAt:yyyy-MM-dd HH:mm}";
            int online = package.Manifest.Software.Count(s => s.RequiresNetwork);
            if (online > 0) heading.Text += $" · {online} 款需要联网";
            foreach (var item in InstallationEngine.Sort(package.Manifest.Software))
            {
                var row = new ListViewItem(item.Name) { Tag = item.Id, ToolTipText = item.Name };
                row.SubItems.Add(item.Version); row.SubItems.Add("待安装");
                row.SubItems.Add(string.Join("；", new[] { item.StartInteractively ? "将打开安装界面" : "", item.RequiresNetwork ? "需要联网" : "" }.Where(s => s.Length > 0))); row.SubItems.Add(item.Category);
                row.SubItems.Add(string.IsNullOrWhiteSpace(item.DefaultInstallDirectory) ? "跟随安装器默认" : item.DefaultInstallDirectory +
                    (item.DirectoryMode == InstallDirectoryMode.Automatic ? "（安装时选择）" : ""));
                rows.Add(item.Id, row);
            }
            categories.Items.Add(new CategoryChoice(null));
            foreach (var name in CategoryNames.Normalize(package.Manifest.Categories.Concat(package.Manifest.Software.Select(s => s.Category))))
                categories.Items.Add(new CategoryChoice(name));
            categories.SelectedIndex = 0;
            schemes.Items.AddRange(package.Manifest.Schemes.Keys.Cast<object>().ToArray());
            schemes.SelectedItem = package.Manifest.DefaultScheme;
            status.Text = "请选择方案并点击一键安装。需要操作时会打开软件安装界面，请按提示完成。";
            install.Enabled = true;
            if (startupSelection != null)
            {
                if (!IsAdministrator()) throw new InvalidOperationException("管理员授权尚未完成。");
                var selected = InstallationEngine.IncludeDependencies(package.Manifest.Software, startupSelection).Select(s => s.Id).ToHashSet();
                SetSelection(selected);
                await BeginInstallAsync();
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
        RenderList();
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
                pair.Value.Checked = selectedIds.Contains(pair.Key);
                if (category == null || pair.Value.SubItems[4].Text.Equals(category, StringComparison.OrdinalIgnoreCase)) list.Items.Add(pair.Value);
            }
        }
        finally { list.EndUpdate(); updatingSelection = false; }
    }

    private static bool IsAdministrator() => new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);

    private async Task BeginInstallAsync()
    {
        if (package == null || running || uncertain) return;
        var selected = selectedIds.ToList();
        if (selected.Count == 0) { status.Text = "请至少勾选一款软件。"; return; }
        var items = InstallationEngine.IncludeDependencies(package.Manifest.Software, selected);
        selectedIds.UnionWith(items.Select(s => s.Id));
        RenderList();
        if (!IsAdministrator())
        {
            try
            {
                Process.Start(new ProcessStartInfo(packagePath, "--install " + string.Join(',', items.Select(s => s.Id))) { UseShellExecute = true, Verb = "runas" });
                Close();
            }
            catch (Exception ex) { status.Text = "未能取得管理员权限：" + ex.Message; }
            return;
        }
        running = true; install.Enabled = retry.Enabled = schemes.Enabled = categories.Enabled = list.Enabled = false; cancel.Enabled = true;
        results = new(); progress.Value = 0;
        cancellation = new();
        var work = Path.Combine(Path.GetTempPath(), "SIA-install-" + Guid.NewGuid().ToString("N"));
        logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SoftwareInstaller", "Logs", DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N")[..8]);
        try
        {
            Directory.CreateDirectory(logDirectory); logs.Enabled = true;
            foreach (ListViewItem row in rows.Values)
                if (items.Any(s => s.Id == (string)row.Tag!)) { row.Checked = true; row.SubItems[2].Text = "准备资源"; row.SubItems[3].Text = ""; }
            var extraction = new Progress<BundleProgress>(p => { status.Text = $"解包及校验资源 {p.Completed}/{p.Total} · {p.Message}"; progress.Style = ProgressBarStyle.Blocks; progress.Value = p.Total == 0 ? 0 : Math.Clamp(p.Completed * 100 / p.Total, 0, 100); });
            await Task.Run(() => package.ExtractAsync(items, work, extraction, cancellation.Token));
            var reporter = new Progress<InstallUpdate>(p =>
            {
                status.Text = $"已处理 {p.Completed}/{p.Total} 款 · {p.StateText} · {p.Message}";
                progress.Style = p.Result == null ? ProgressBarStyle.Marquee : ProgressBarStyle.Blocks;
                if (p.Result != null) progress.Value = p.Total == 0 ? 0 : p.Completed * 100 / p.Total;
                var row = rows[p.Id];
                row.SubItems[2].Text = p.StateText;
                row.SubItems[3].Text = p.Message;
            });
            results = await Task.Run(() => InstallationEngine.RunAsync(items, work, logDirectory, reporter, cancellation.Token));
            uncertain = results.Any(r => r.State == InstallState.Uncertain || r.ExitCode == 1641);
            var finished = results.Select(r => r.Id).ToHashSet();
            foreach (ListViewItem row in rows.Values)
                if (items.Any(i => i.Id == (string)row.Tag!) && !finished.Contains((string)row.Tag!)) { row.SubItems[2].Text = "未执行"; row.SubItems[3].Text = "队列已停止"; }
            status.Text = $"已处理 {results.Count}/{items.Count} 项：完成 {results.Count(r => r.State == InstallState.Succeeded)}，跳过 {results.Count(r => r.State == InstallState.Skipped)}，失败 {results.Count(r => r.State == InstallState.Failed)}，待检查 {results.Count(r => r.State == InstallState.Uncertain)}，需重启 {results.Count(r => r.State == InstallState.RebootRequired)}，未执行 {items.Count - results.Count}。";
            if (results.Any(r => r.State == InstallState.RebootRequired)) status.Text += " 部分软件需要重启。";
            if (uncertain) { retained.Add(work); status.Text += " 队列已暂停；请检查仍在运行的安装器，资源暂时保留。"; }
        }
        catch (OperationCanceledException) { status.Text = "已取消准备安装。"; }
        catch (Exception ex)
        {
            // Preserve files if an unexpected failure occurs after an installer may have started.
            retained.Add(work); uncertain = true;
            status.Text = "安装流程已停止：" + ex.Message;
            try { await File.WriteAllTextAsync(Path.Combine(logDirectory, "error.log"), ex.ToString()); }
            catch (Exception logError) when (logError is IOException or UnauthorizedAccessException) { status.Text += " 日志无法写入。"; }
        }
        finally
        {
            progress.Style = ProgressBarStyle.Blocks;
            progress.Value = items.Count == 0 ? 0 : results.Count * 100 / items.Count;
            if (!retained.Contains(work) && Directory.Exists(work))
            {
                try { Directory.Delete(work, true); }
                catch (IOException) { retained.Add(work); }
                catch (UnauthorizedAccessException) { retained.Add(work); }
            }
            if (retained.Count > 0)
            {
                try { await File.WriteAllLinesAsync(Path.Combine(logDirectory, "retained-resources.txt"), retained); }
                catch (Exception logError) when (logError is IOException or UnauthorizedAccessException) { status.Text += " 资源目录：" + work; }
                status.Text += " 保留目录见日志。";
            }
            running = false; cancellation.Dispose(); cancellation = null; cancel.Enabled = false;
            install.Enabled = schemes.Enabled = categories.Enabled = list.Enabled = !uncertain;
            retry.Enabled = !uncertain && results.Any(r => r.State == InstallState.Failed);
        }
    }

}
