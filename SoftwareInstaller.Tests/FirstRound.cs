using System.Diagnostics;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using SoftwareInstaller.Core;
using SoftwareInstaller.Models;
using SoftwareInstaller.Packaging;
using SoftwareInstaller.UI;

namespace SoftwareInstaller.Tests;

internal static partial class Program
{
    private sealed class Updates(List<InstallUpdate> items) : IProgress<InstallUpdate>
    {
        public void Report(InstallUpdate update) => items.Add(update);
    }

    private static async Task FirstRoundAsync()
    {
        var exe = Path.Combine(root, "办公助手.exe"); var msi = Path.Combine(root, "开发工具.msi");
        File.WriteAllBytes(exe, [77, 90, 0, 1]); File.WriteAllBytes(msi, [77, 90, 0, 2]);
        var preparation = SoftwareImport.Prepare([exe, msi, exe, Path.Combine(root, "not-installer.txt")], [], "办公");
        Assert(preparation.Items.Count == 2 && preparation.Duplicates == 1 && preparation.Errors.Count == 1, "batch import separates installers, duplicate paths and invalid files");
        Assert(preparation.Items[0].Name == "办公助手" && preparation.Items[1].Name == "开发工具", "EXE and MSI filenames fill batch names without running installers");
        Assert(SoftwareImport.Prepare([exe], preparation.Items, "办公").Duplicates == 1, "already configured installer is reported as a duplicate");
        PreviewFirstRound(preparation);
        SynchronizationContext.SetSynchronizationContext(null);

        var invalidRunner = Path.Combine(root, "old-template.exe"); File.WriteAllBytes(invalidRunner, [77, 90, 0]);
        bool rejected = false;
        try { RunnerTemplate.Validate(invalidRunner); } catch (InvalidDataException) { rejected = true; }
        Assert(rejected, "an unversioned or obsolete runner cannot silently generate packages");
        RunnerTemplate.Validate(runner);
        Assert(SHA256.HashData(File.ReadAllBytes(runner)).SequenceEqual(SHA256.HashData(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Runner", "SoftwareInstaller.Runner.exe")))),
            "management preview uses the newly built runner template");
        var output = Path.Combine(root, "高清图标验证合集.exe");
        await BundleBuilder.BuildAsync(new() { Name = "安装器界面预览", RunnerPath = runner, OutputPath = output,
            Software = preparation.Items, Categories = ["办公"], DefaultScheme = "全部软件", Schemes = new() { ["全部软件"] = preparation.Items.Select(s => s.Id).ToList() } });
        var report = Path.Combine(root, "export-report.json"); var preview = Path.Combine(root, "preview-exported-installer.png");
        var info = new ProcessStartInfo(output) { UseShellExecute = false, CreateNoWindow = true };
        foreach (var argument in new[] { "--verify-package", report, "--preview", preview }) info.ArgumentList.Add(argument);
        using (var process = Process.Start(info)!)
        {
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(45)); await process.WaitForExitAsync(timeout.Token);
            Assert(process.ExitCode == 0 && File.Exists(preview), "the actual exported EXE starts, verifies resources and renders its own installer window");
        }
        using var document = JsonDocument.Parse(File.ReadAllText(report));
        Assert(document.RootElement.GetProperty("runnerVersion").GetString() == RunnerTemplate.CurrentVersion && document.RootElement.GetProperty("logoWidth").GetInt32() >= 1000,
            "the exported EXE contains the current runner and full-resolution PNG logo");

        var safe = Path.Combine(root, "safe-fallback"); Directory.CreateDirectory(safe);
        foreach (var path in Directory.EnumerateFiles(AppContext.BaseDirectory).Where(p => !p.EndsWith(".pdb"))) File.Copy(path, Path.Combine(safe, Path.GetFileName(path)), true);
        var helper = Path.Combine(safe, "FallbackFixture.exe"); File.Copy(Environment.ProcessPath!, helper);
        var ledger = Path.Combine(safe, "attempts.txt"); File.WriteAllText(Path.Combine(safe, "scenario.txt"), ledger);
        var updates = new List<InstallUpdate>();
        var results = await InstallationEngine.RunAsync([new() { Id = Guid.NewGuid().ToString("N"), Name = "无副作用安装器替身", Installer = "FallbackFixture.exe", Arguments = "--reject", DetectionFile = ledger }],
            safe, Path.Combine(root, "phase-logs"), new Updates(updates), resolveInstaller: _ => helper);
        Assert(updates.Select(p => p.Phase).ToHashSet().IsSupersetOf([InstallPhase.Checking, InstallPhase.Installing, InstallPhase.WaitingForUser, InstallPhase.Verifying]) && results.Single().State == InstallState.Succeeded,
            "real safe child processes report checking, installation, interactive fallback, verification and completion phases");
    }

    private static void PreviewFirstRound(ImportPreparation preparation)
    {
        Application.EnableVisualStyles();
        using var main = new MainForm { ShowInTaskbar = false, Opacity = 0 };
        T Field<T>(object form, string name) => (T)form.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(form)!;
        var manager = Field<SoftwareManager>(main, "_softwareManager");
        preparation.Items[1].Category = "开发"; preparation.Items[1].IsSelected = false;
        foreach (var item in preparation.Items) manager.AddSoftwareToActiveList(item.Clone());
        typeof(MainForm).GetMethod("RebuildTreeView", BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(main, null);
        main.Show(); Application.DoEvents();
        var list = Field<ListView>(main, "softwareListView");
        Assert(list.Items.Count == 2 && list.CheckedItems.Count == 1, "unchecked software remains in the complete management list");
        list.Items[0].Checked = false; Application.DoEvents();
        Assert(list.Items.Count == 2 && manager.ActiveSoftwareItems.All(s => !s.IsSelected), "unchecking a row updates the model without hiding the row");
        list.Items[1].Checked = true;
        var search = Field<TextBox>(main, "softwareSearch"); search.Text = "办公"; Application.DoEvents();
        Assert(list.Items.Count == 1 && manager.ActiveSoftwareItems[1].IsSelected, "search filters unchecked rows while preserving hidden selections");
        search.Clear(); var category = Field<ComboBox>(main, "softwareCategory"); category.SelectedIndex = 1; Application.DoEvents();
        Assert(list.Items.Count == 1 && manager.ActiveSoftwareItems[1].IsSelected, "category filtering preserves selections outside the visible category");
        category.SelectedIndex = 0;
        using (var bitmap = new Bitmap(main.Width, main.Height)) { main.DrawToBitmap(bitmap, new Rectangle(Point.Empty, main.Size)); bitmap.Save(Path.Combine(root, "preview-management.png")); }
        var drafts = new ImportPreparation(preparation.Items.Select(s => { var copy = s.Clone(); copy.IsSelected = true; return copy; }).ToList(), 1, []);
        using var batch = new BatchImportForm(drafts, ["办公", "开发"]) { ShowInTaskbar = false, Opacity = 0 };
        batch.Show(); Application.DoEvents();
        var grid = Field<DataGridView>(batch, "grid"); grid.Rows[0].Cells[1].Value = "自定义名称"; grid.EndEdit();
        Assert(batch.Software[0].Name == "自定义名称" && batch.Software.Count == 2, "batch confirmation retains editable names and multiple installers");
        using var picture = new Bitmap(batch.Width, batch.Height); batch.DrawToBitmap(picture, new Rectangle(Point.Empty, batch.Size)); picture.Save(Path.Combine(root, "preview-batch.png"));
    }
}
