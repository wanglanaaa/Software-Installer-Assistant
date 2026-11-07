using System.Diagnostics;
using System.IO.Compression;
using System.Reflection;
using System.Text.Json;
using SoftwareInstaller.Packaging;
using SoftwareInstaller.Runner;

namespace SoftwareInstaller.Tests;

internal static partial class Program
{
    private static async Task TestRunnerPerformanceAsync()
    {
        var prepared = new List<string>();
        var completed = new List<string>();
        var markers = Path.Combine(root, "pipeline"); Directory.CreateDirectory(markers);
        BundleSoftware Item(string name) => new()
        {
            Id = Guid.NewGuid().ToString("N"), Name = name, Installer = "setup.exe", AllowInteractiveFallback = false,
            Arguments = $"--process-fixture sleep 30 \"{Path.Combine(markers, name)}\""
        };
        Task<List<InstallResult>> Run(BundleSoftware[] items, string name, Func<BundleSoftware, CancellationToken, Task> prepare,
            CancellationToken token = default, InstallationControl? control = null) => InstallationEngine.RunAsync(InstallationPlan.Create(items), root,
            Path.Combine(root, name), new InlineProgress(p => { if (p.Result != null) completed.Add(p.Id); }), token,
            resolveInstaller: _ => Environment.ProcessPath!, control: control, prepareInstaller: prepare);

        var first = Item("first"); var second = Item("second"); second.DependencyIds.Add(first.Id);
        var already = Item("already"); already.DetectionFile = Path.Combine(root, "already-installed"); File.WriteAllText(already.DetectionFile, "installed");
        var source = new List<BundleSoftware> { second, already, first };
        var plan = InstallationPlan.Create(source);
        source.Clear();
        Assert(plan.Select([second.Id, second.Id]).Items.Select(s => s.Id).SequenceEqual([first.Id, second.Id])
            && plan.Select([already.Id]).Items.Single().Id == already.Id && plan.Items.Count == 3,
            "a reusable plan preserves dependency order and independent selections after the source list changes");
        await Reject(() => { plan.Select([Guid.NewGuid().ToString("N")]); return Task.CompletedTask; }, "cached selection still rejects unknown software IDs");
        await Reject(() => { InstallationPlan.Create([second]); return Task.CompletedTask; }, "plan creation still rejects missing dependencies");
        var cycleA = Item("cycle-a"); var cycleB = Item("cycle-b"); cycleA.DependencyIds.Add(cycleB.Id); cycleB.DependencyIds.Add(cycleA.Id);
        await Reject(() => { InstallationPlan.Create([cycleA, cycleB]); return Task.CompletedTask; }, "plan creation still rejects dependency cycles");
        var results = await Run([second, already, first], "logs-pipeline", (item, _) =>
        {
            if (item.Id == second.Id)
                Assert(completed.Contains(first.Id) && File.Exists(Path.Combine(markers, "first")), "the first installer finishes before preparing the next software's resources");
            prepared.Add(item.Id);
            return Task.CompletedTask;
        });
        Assert(prepared.SequenceEqual([first.Id, second.Id]) && results.Single(r => r.Id == already.Id).State == InstallState.Skipped,
            "installed software is skipped without unpacking and dependency order is preserved");
        Assert(results.Where(r => r.State != InstallState.Skipped).All(r => r.State == InstallState.Succeeded), "per-software preparation still launches safe fixture installers successfully");

        prepared.Clear();
        var failed = Item("failed-preparation"); var blocked = Item("blocked"); blocked.DependencyIds.Add(failed.Id);
        results = await Run([failed, blocked], "logs-preparation-failure", (item, _) =>
        { prepared.Add(item.Id); throw new IOException("Synthetic extraction failure"); });
        Assert(prepared.SequenceEqual([failed.Id]) && results.All(r => r.State == InstallState.Failed) && !File.Exists(Path.Combine(markers, "blocked")),
            "failed preparation prevents installation and does not unpack dependent software");

        using (var stop = new CancellationTokenSource())
        {
            var cancelled = Item("cancelled-preparation");
            results = await Run([cancelled, Item("after-cancel")], "logs-preparation-cancel", (_, token) =>
            { stop.Cancel(); token.ThrowIfCancellationRequested(); return Task.CompletedTask; }, token: stop.Token);
            Assert(results.Single().State == InstallState.Cancelled && !File.Exists(Path.Combine(markers, "cancelled-preparation")),
                "stopping during unpacking cancels preparation without starting an installer or the next item");
        }
        using (var control = new InstallationControl())
        {
            var waiting = Item("stop-wait-preparation");
            results = await Run([waiting], "logs-preparation-stop-wait", (_, token) =>
            { control.StopWaiting(); token.ThrowIfCancellationRequested(); return Task.CompletedTask; }, control: control);
            Assert(results.Single().State == InstallState.Cancelled && !File.Exists(Path.Combine(markers, "stop-wait-preparation")),
                "stop-wait requests also interrupt resource preparation before installation");
        }

        // A UI-only archive exercises batched loading without a native launcher or executable resources.
        var resource = new BundleFile { Path = "setup.exe", Sha256 = new string('0', 64), Size = 0 };
        var software = Enumerable.Range(1, 120).Select(i => new BundleSoftware { Id = Guid.NewGuid().ToString("N"),
            Name = $"演示软件 {i:000}", Category = i % 2 == 1 ? "演示分类" : "其他分类", Installer = resource.Path, Files = [resource] }).ToList();
        var manifest = new BundleManifest { Name = "合集界面预览", Software = software,
            Schemes = new() { ["当前方案"] = software.Select(s => s.Id).ToList(), ["部分方案"] = software.Take(3).Select(s => s.Id).ToList() } };
        using var payload = new MemoryStream();
        using (var archive = new ZipArchive(payload, ZipArchiveMode.Create, true))
        {
            archive.CreateEntry("content/" + resource.Sha256);
            await using var metadata = archive.CreateEntry("manifest.json").Open();
            await JsonSerializer.SerializeAsync(metadata, manifest);
        }
        await using var output = File.Create(Path.Combine(root, "ui-preview.bundle"));
        await output.WriteAsync(payload.ToArray()); await output.WriteAsync(BitConverter.GetBytes(payload.Length));
        await output.WriteAsync(new byte[32]); await output.WriteAsync(BundleReader.Magic);
    }

    private static void RenderRunnerPerformancePreview()
    {
        Application.EnableVisualStyles();
        using var form = new InstallForm(Path.Combine(root, "ui-preview.bundle")) { ShowInTaskbar = false, Opacity = 0 };
        form.Show();
        T Field<T>(string name) => (T)typeof(InstallForm).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
        var watch = Stopwatch.StartNew();
        while (!Field<Button>("install").Enabled && watch.Elapsed < TimeSpan.FromSeconds(10)) { Application.DoEvents(); Thread.Sleep(1); }
        var list = Field<ListView>("list");
        var schemes = Field<ComboBox>("schemes"); var categories = Field<ComboBox>("categories");
        var selectCategory = Field<CheckBox>("selectCategory");
        void ToggleCategory() => typeof(CheckBox).GetMethod("OnClick", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(selectCategory, [EventArgs.Empty]);
        Assert(Field<Button>("install").Enabled && list.Items.Count == 120 && list.CheckedItems.Count == 120,
            "the runner loads and selects a large software list through the batched UI path");
        var selectedRow = list.Items[60]; selectedRow.Selected = true; list.EnsureVisible(60); Application.DoEvents();
        int top = list.TopItem!.Index;
        ToggleCategory(); Application.DoEvents();
        Assert(list.CheckedItems.Count == 0 && selectedRow.Selected && list.TopItem!.Index == top,
            "bulk checkbox updates preserve the highlighted software and scroll position");
        schemes.SelectedItem = "部分方案"; Application.DoEvents();
        Assert(list.CheckedItems.Count == 3 && selectedRow.Selected && list.TopItem!.Index == top,
            "scheme changes synchronize checkboxes without rebuilding the visible list");
        string hiddenId = (string)list.Items[1].Tag!;
        categories.SelectedIndex = 1; Application.DoEvents();
        ToggleCategory();
        Assert(list.Items.Count == 60 && list.CheckedItems.Count == 60 && Field<HashSet<string>>("selectedIds").Contains(hiddenId),
            "category selection preserves checked software hidden in other categories");
        categories.SelectedIndex = 0;
        Assert(list.CheckedItems.Count == 61, "returning to all categories restores synchronized hidden checkboxes");
        schemes.SelectedItem = "当前方案";
        list.Items[0].Checked = false;
        Assert(list.CheckedItems.Count == 119 && selectCategory.CheckState == CheckState.Indeterminate,
            "individual checkbox changes still update the category's mixed selection state");
        list.Items[0].Checked = true; list.EnsureVisible(0);
        foreach (var (name, size) in new[] { ("default", new Size(920, 650)), ("minimum", new Size(760, 500)) })
        {
            form.Size = size; Application.DoEvents();
            using var bitmap = new Bitmap(form.Width, form.Height);
            form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.Size)); bitmap.Save(Path.Combine(root, "preview-runner-" + name + ".png"));
        }
        form.Close();
    }
}
