using System.Diagnostics;
using SoftwareInstaller.Core;
using SoftwareInstaller.Models;
using SoftwareInstaller.Packaging;

namespace SoftwareInstaller.UI;

public sealed class BundleBuildForm : Form
{
    private readonly List<SoftwareItem> current;
    private readonly Dictionary<string, List<SoftwareItem>> saved;
    private readonly List<string> currentCategories;
    private readonly Dictionary<string, List<string>> savedCategories;
    private readonly TabControl pages = new() { Dock = DockStyle.Fill };
    private readonly CheckBox onlyChecked = new() { Text = "仅打包当前勾选软件（自动加入前置软件）", AutoSize = true };
    private readonly CheckedListBox schemes = new() { Dock = DockStyle.Fill, CheckOnClick = true };
    private readonly ListView checks = new() { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true };
    private readonly TextBox name = new() { Text = "我的软件合集", Dock = DockStyle.Top };
    private readonly TextBox output = new() { Dock = DockStyle.Top, ReadOnly = true };
    private readonly ComboBox defaultScheme = new() { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
    private readonly Label status = new() { Dock = DockStyle.Bottom, Height = 48, Text = "选择内容后点击下一步。" };
    private readonly ProgressBar progress = new() { Dock = DockStyle.Bottom, Height = 16 };
    private readonly Button previous = new() { Text = "上一步", AutoSize = true }, next = new() { Text = "下一步", AutoSize = true }, cancel = new() { Text = "关闭", AutoSize = true };
    private CancellationTokenSource? cancellation;
    private BundleBuildRequest? prepared;
    private bool busy, checkedReady;

    public BundleBuildForm(IEnumerable<SoftwareItem> software, SchemeHandler handler, IEnumerable<string>? categories = null)
    {
        AppBranding.Apply(this);
        current = software.Select(s => s.Clone()).ToList();
        currentCategories = CategoryNames.Normalize((categories ?? []).Concat(current.Select(s => s.Category)));
        saved = handler.GetSchemeNames().ToDictionary(n => n, n => handler.GetScheme(n)!);
        savedCategories = handler.GetSchemeNames().ToDictionary(n => n, handler.GetCategories);
        Text = "生成单文件安装包"; Size = new(840, 640); MinimumSize = new(720, 520); StartPosition = FormStartPosition.CenterParent;
        var content = new TabPage("① 选择内容") { Padding = new(16) };
        var help = new Label { Text = $"当前列表 {current.Count} 款，已勾选 {current.Count(s => s.IsSelected)} 款，分类 {currentCategories.Count} 个（含空分类）。\n默认携带全部软件；没有勾选时自动创建「全部软件」方案。\n附带方案中选中的软件及前置软件会一起打包；相同软件使用当前列表配置。", Height = 90, Dock = DockStyle.Top };
        var choices = new Panel { Dock = DockStyle.Top, Height = 64 };
        onlyChecked.Dock = DockStyle.Top;
        choices.Controls.Add(new Label { Text = "可选：附带已保存方案", Dock = DockStyle.Bottom, Height = 24 }); choices.Controls.Add(onlyChecked);
        foreach (var key in saved.Keys) schemes.Items.Add(key);
        content.Controls.Add(schemes); content.Controls.Add(choices); content.Controls.Add(help);
        var checkPage = new TabPage("② 检查配置") { Padding = new(16) };
        checks.Columns.Add("软件", 210); checks.Columns.Add("检查结果", 480);
        checkPage.Controls.Add(checks);
        var refresh = new Button { Text = "重新检查", Dock = DockStyle.Bottom, Height = 34 };
        refresh.Click += async (_, _) => await CheckAsync(); checkPage.Controls.Add(refresh);
        var generate = new TabPage("③ 生成文件") { Padding = new(24) };
        var fields = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1 };
        void Field(string label, Control control) { fields.Controls.Add(new Label { Text = label, AutoSize = true, Margin = new(0, 12, 0, 4) }); control.Width = 600; control.Margin = new(0, 0, 0, 8); fields.Controls.Add(control); }
        Field("包名称", name); Field("默认安装方案", defaultScheme); Field("输出 EXE", output);
        var browse = new Button { Text = "选择输出位置…", AutoSize = true };
        browse.Click += (_, _) => { using var dialog = new SaveFileDialog { Filter = "EXE 安装包|*.exe", FileName = "软件合集.exe", OverwritePrompt = false }; if (dialog.ShowDialog(this) == DialogResult.OK) output.Text = dialog.FileName; };
        fields.Controls.Add(browse);
        fields.Controls.Add(new Label { Text = "生成的安装包适用于 Windows x64，内置运行时。\n单文件上限约 3.4 GiB；联网安装器需要目标电脑联网。\n无法自动安装时，会打开软件原有安装界面。", AutoSize = true, Margin = new(0, 24, 0, 0) });
        generate.Controls.Add(fields);
        pages.TabPages.AddRange([content, checkPage, generate]);
        onlyChecked.CheckedChanged += (_, _) => ResetChecks();
        schemes.ItemCheck += (_, _) => ResetChecks();
        pages.SelectedIndexChanged += (_, _) =>
        {
            previous.Enabled = pages.SelectedIndex > 0; next.Text = pages.SelectedIndex == 2 ? "生成安装包" : "下一步";
            if (pages.SelectedIndex == 0) ResetChecks();
        };
        // Navigation is deliberate, so unchecked content cannot bypass validation.
        pages.Selecting += (_, e) => { if (!busy && e.TabPageIndex > pages.SelectedIndex) e.Cancel = true; };
        previous.Click += (_, _) => { if (pages.SelectedIndex > 0) SetPage(pages.SelectedIndex - 1); };
        next.Click += async (_, _) =>
        {
            if (pages.SelectedIndex == 0) { SetPage(1); await CheckAsync(); }
            else if (pages.SelectedIndex == 1) { if (checkedReady) SetPage(2); else await CheckAsync(); }
            else await BuildAsync();
        };
        cancel.Click += (_, _) => { if (busy) cancellation?.Cancel(); else Close(); };
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new(12) };
        buttons.Controls.Add(cancel); buttons.Controls.Add(next); buttons.Controls.Add(previous);
        Controls.Add(pages); Controls.Add(status); Controls.Add(progress); Controls.Add(buttons);
        FormClosing += (_, e) => { if (busy) { e.Cancel = true; cancellation?.Cancel(); status.Text = "正在取消，请等待资源写入停止…"; } };
        SetPage(0);
    }

    private void SetPage(int index)
    {
        busy = true; pages.SelectedIndex = index; busy = false;
        previous.Enabled = index > 0; next.Text = index == 2 ? "生成安装包" : "下一步";
        if (index == 0) ResetChecks();
    }

    private void ResetChecks()
    {
        checkedReady = false; prepared = null; checks.Items.Clear(); defaultScheme.Items.Clear(); progress.Value = 0;
        status.Text = "选择内容后点击下一步，将重新检查当前配置。";
    }

    private BundleBuildRequest Prepare()
    {
        var all = current.ToDictionary(s => s.Id);
        foreach (var key in schemes.CheckedItems.Cast<string>())
            foreach (var item in saved[key]) all.TryAdd(item.Id, item);
        var included = new Dictionary<string, SoftwareItem>();
        void Add(string id)
        {
            if (included.ContainsKey(id)) return;
            if (!all.TryGetValue(id, out var item)) throw new InvalidDataException("方案缺少前置软件，请重新保存方案。");
            included.Add(id, item);
            foreach (var dependency in item.DependencyIds) Add(dependency);
        }
        foreach (var item in current.Where(s => !onlyChecked.Checked || s.IsSelected)) Add(item.Id);
        var plans = new Dictionary<string, List<string>>();
        var selected = current.Where(s => s.IsSelected).Select(s => s.Id).ToList();
        if (selected.Count > 0) plans.Add("当前方案", selected);
        foreach (var key in schemes.CheckedItems.Cast<string>())
        {
            var ids = saved[key].Where(s => s.IsSelected).Select(s => s.Id).Distinct().ToList();
            if (ids.Count == 0) throw new InvalidDataException($"方案「{key}」没有勾选软件。");
            foreach (var id in ids) Add(id);
            plans.Add("已保存：" + key, ids);
        }
        if (plans.Count == 0 && !onlyChecked.Checked && included.Count > 0)
            plans.Add("全部软件", included.Keys.ToList());
        if (plans.Count == 0) throw new InvalidDataException(onlyChecked.Checked
            ? "未勾选软件。请返回选择「打包全部」，或在主界面勾选软件。"
            : "当前配置没有软件。请先加载配置文件或新增软件。");
        return new()
        {
            Software = included.Values.Select(s => s.Clone()).ToList(),
            Schemes = plans,
            DefaultScheme = plans.Keys.First(),
            RunnerPath = FindRunner(),
            Categories = CategoryNames.Normalize(currentCategories.Concat(schemes.CheckedItems.Cast<string>().SelectMany(s => savedCategories[s])))
        };
    }

    private static string FindRunner()
    {
        var installed = Path.Combine(AppContext.BaseDirectory, "Runner", "SoftwareInstaller.Runner.exe");
        if (File.Exists(installed)) return installed;
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory != null; directory = directory.Parent)
        {
            var path = Path.Combine(directory.FullName, "artifacts", "runner", "SoftwareInstaller.Runner.exe");
            if (File.Exists(path)) return path;
        }
        return installed;
    }

    private async Task CheckAsync()
    {
        ResetChecks();
        busy = true; next.Enabled = previous.Enabled = false; pages.Enabled = false; status.Text = "正在检查当前配置…";
        try
        {
            prepared = Prepare(); bool ready = true;
            foreach (var item in prepared.Software)
            {
                var errors = await Task.Run(() => BundleBuilder.CheckItem(item));
                ready &= errors.Count == 0;
                var row = checks.Items.Add(item.Name);
                var description = errors.Count == 0 ? await Task.Run(() => BundleBuilder.Describe(item)) : null;
                row.SubItems.Add(errors.Count == 0 ? "可打包 · " + (description!.StartInteractively ? "普通安装" : "优先自动安装，失败时打开界面") +
                    (item.RequiresNetwork ? "；需要联网" : "") : string.Join("；", errors));
                row.ForeColor = errors.Count == 0 ? Color.DarkGreen : Color.Firebrick;
            }
            if (!File.Exists(prepared.RunnerPath)) { ready = false; checks.Items.Add("执行器").SubItems.Add("未找到自包含执行器，请先运行 scripts/publish-portable.ps1"); }
            InstallationEngine.Sort(prepared.Software.Select(s => new BundleSoftware { Id = s.Id, Name = s.Name, DependencyIds = s.DependencyIds, Order = s.InstallOrder }));
            defaultScheme.Items.Clear(); defaultScheme.Items.AddRange(prepared.Schemes.Keys.Cast<object>().ToArray()); defaultScheme.SelectedIndex = 0;
            checkedReady = ready;
            status.Text = ready ? $"检查通过：{prepared.Software.Count} 款软件，{prepared.Schemes.Count} 个方案。" : "请关闭此窗口，编辑对应软件后重新生成。";
        }
        catch (Exception ex) { status.Text = ex.Message; }
        finally { busy = false; next.Enabled = previous.Enabled = true; pages.Enabled = true; }
    }

    private async Task BuildAsync()
    {
        if (prepared == null || !checkedReady) return;
        if (string.IsNullOrWhiteSpace(output.Text)) { status.Text = "请选择输出 EXE 的位置。"; return; }
        if (File.Exists(output.Text) && MessageBox.Show(this, "输出文件已存在，是否替换？", "覆盖文件", MessageBoxButtons.YesNo) != DialogResult.Yes) return;
        var request = new BundleBuildRequest
        {
            Name = name.Text,
            OutputPath = output.Text,
            RunnerPath = prepared.RunnerPath,
            Software = prepared.Software,
            Categories = prepared.Categories,
            Schemes = prepared.Schemes,
            DefaultScheme = (string)defaultScheme.SelectedItem!
        };
        busy = true; next.Enabled = previous.Enabled = pages.Enabled = false; cancel.Text = "取消生成"; cancellation = new();
        try
        {
            var reporter = new Progress<BundleProgress>(p => { status.Text = p.Message; progress.Value = p.Total == 0 ? 0 : Math.Clamp(p.Completed * 100 / p.Total, 0, 100); });
            await Task.Run(() => BundleBuilder.BuildAsync(request, reporter, cancellation.Token));
            progress.Value = 100; status.Text = "生成完成：" + request.OutputPath;
            try { Process.Start(new ProcessStartInfo("explorer.exe", $"/select,\"{request.OutputPath}\"") { UseShellExecute = true }); }
            catch (System.ComponentModel.Win32Exception) { /* The generated file remains available at the displayed path. */ }
        }
        catch (OperationCanceledException) { status.Text = "已取消，原有输出文件保留。"; }
        catch (Exception ex) { status.Text = "生成失败：" + ex.Message; }
        finally { busy = false; cancellation.Dispose(); cancellation = null; next.Enabled = previous.Enabled = pages.Enabled = true; cancel.Text = "关闭"; }
    }
}
