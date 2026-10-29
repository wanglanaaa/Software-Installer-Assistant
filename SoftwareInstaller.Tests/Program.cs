using System.Diagnostics;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using System.Runtime.InteropServices;
using SoftwareInstaller.Core;
using SoftwareInstaller.Models;
using SoftwareInstaller.Packaging;
using SoftwareInstaller.UI;
using SoftwareInstaller.Runner;

namespace SoftwareInstaller.Tests;

internal static class Program
{
    private static int passed;
    private static string root = "";
    private static string runner = "";
    private static string fixture = "";

    [STAThread]
    private static int Main(string[] args)
    {
        if (string.Equals(Path.GetFileName(Environment.ProcessPath), "FallbackFixture.exe", StringComparison.OrdinalIgnoreCase))
        {
            var ledger = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "scenario.txt"));
            File.AppendAllText(ledger, args.Length == 0 ? "interactive\n" : "silent\n");
            return args.Length == 0 ? 0 : args[0] == "--cancel" ? 1602 : 1639;
        }
        if (args.Length >= 7 && args[0] == "--fake")
        {
            Thread.Sleep(int.Parse(args[6]));
            File.AppendAllText(args[2], args[3] + Environment.NewLine);
            if (args[4] != "-") File.WriteAllText(args[4], "installed");
            int directoryIndex = Array.IndexOf(args, "--target-dir");
            if (directoryIndex >= 0 && directoryIndex + 1 < args.Length)
            {
                Directory.CreateDirectory(args[directoryIndex + 1]);
                File.WriteAllText(Path.Combine(args[directoryIndex + 1], "installed.marker"), "installed");
                File.WriteAllText(Path.Combine(args[directoryIndex + 1], "received.json"), JsonSerializer.Serialize(args));
            }
            return int.Parse(args[1]);
        }
        try
        {
            root = Path.GetFullPath(Path.Combine(".build", "integration-" + Guid.NewGuid().ToString("N")));
            runner = Path.GetFullPath("artifacts/runner/SoftwareInstaller.Runner.exe");
            Directory.CreateDirectory(root);
            UserDataPaths.DataDirectoryOverride = Path.Combine(root, "user-data");
            if (args.Contains("--branding-only"))
            {
                TestNameAndBranding();
                Console.WriteLine($"PASS: {passed} name and branding checks. Preview: {root}"); return 0;
            }
            if (args.Contains("--configuration-only"))
            {
                TestManualConfiguration();
                Console.WriteLine($"PASS: {passed} configuration workflow checks. Preview: {root}"); return 0;
            }
            TestCategoryAndDirectorySettings();
            if (args.Contains("--settings-only")) { Console.WriteLine($"PASS: {passed} settings checks."); return 0; }
            RenderPreviews();
            if (args.Contains("--ui-only")) { Console.WriteLine($"PASS: {passed} checks. UI artifacts: {root}"); return 0; }
            SynchronizationContext.SetSynchronizationContext(null);
            RunAsync().GetAwaiter().GetResult();
            RenderRunnerAndWizard();
            Console.WriteLine($"PASS: {passed} checks. Artifacts: {root}");
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new Exception("FAIL: " + message);
        passed++; Console.WriteLine("PASS: " + message);
    }

    private static void TestNameAndBranding()
    {
        Application.EnableVisualStyles();
        var path = Path.Combine(root, "办公工具.v2-x64.msi");
        var item = new SoftwareItem { Name = "", Version = "", Size = "", Description = "", Category = "未分类", FilePath = path };
        using var editor = new SoftwareEditForm(item);
        TextBox Field(string name) => (TextBox)typeof(SoftwareEditForm).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(editor)!;
        var name = Field("nameTextBox"); var file = Field("filePathTextBox");
        Assert(name.Text == "办公工具.v2-x64", "MSI filename fills an empty software name while keeping dots and Chinese characters");
        file.Text = Path.Combine(root, "Setup助手-1.0.exe");
        Assert(name.Text == "Setup助手-1.0", "changing installers updates the automatically filled filename");
        editor.Text = "新增软件"; editor.ShowInTaskbar = false; editor.Opacity = 0; editor.Show(); Application.DoEvents();
        using (var bitmap = new Bitmap(editor.Width, editor.Height)) { editor.DrawToBitmap(bitmap, new Rectangle(Point.Empty, editor.Size)); bitmap.Save(Path.Combine(root, "preview-name-icon.png")); }
        name.Text = "我自定义的名称"; file.Text = Path.Combine(root, "other.exe");
        Assert(name.Text == "我自定义的名称", "selecting another installer preserves a manually edited name");
        item.Name = "原有的软件名";
        using var existing = new SoftwareEditForm(item);
        Assert(((TextBox)typeof(SoftwareEditForm).GetField("nameTextBox", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(existing)!).Text == "原有的软件名",
            "opening an existing software entry preserves its configured name");
        using var main = new MainForm();
        using var wizard = new BundleBuildForm([], new SchemeHandler(useDefaultStorage: false));
        using var input = new InputDialog("新增分类", "分类名");
        using var runnerForm = new InstallForm(Path.Combine(root, "not-started.exe"));
        Assert(new Form[] { main, editor, existing, wizard, input, runnerForm }.All(f => f.ShowIcon && f.Icon?.Size == new Size(32, 32)),
            "main, editor, wizard, input and installer windows load the embedded application icon");
        using var exeIcon = Icon.ExtractAssociatedIcon(Path.Combine(AppContext.BaseDirectory, "SoftwareInstaller.UI.exe"));
        using var exeBitmap = exeIcon?.ToBitmap(); using var windowBitmap = main.Icon!.ToBitmap();
        Assert(exeBitmap != null && exeBitmap.GetPixel(16, 16).ToArgb() == windowBitmap.GetPixel(16, 16).ToArgb(),
            "the executable file uses the same icon as its window");
        main.ShowInTaskbar = false; main.Opacity = 0; main.Show(); Application.DoEvents();
        using (var bitmap = new Bitmap(main.Width, main.Height)) { main.DrawToBitmap(bitmap, new Rectangle(Point.Empty, main.Size)); bitmap.Save(Path.Combine(root, "preview-main-icon.png")); }
    }

    private static void TestManualConfiguration()
    {
        // Old default data must neither affect a new session nor be overwritten by it.
        UserDataPaths.Write("categories.json", "[\"旧分类\"]");
        UserDataPaths.Write("schemes.json", "invalid old JSON");
        UserDataPaths.Write("software_list.json", "invalid old JSON");
        var source = Path.Combine(root, "setup.exe"); File.WriteAllBytes(source, [77, 90, 0]);
        var item = new SoftwareItem { Name = "新软件", Version = "1", Size = "", Description = "", Category = "新分类", FilePath = source, DefaultInstallDirectory = "D:\\" };
        var configuration = new WorkspaceConfiguration { Software = [item], Categories = ["新分类", "空分类"],
            Schemes = new() { ["常用"] = new() { Software = [item.Clone()], Categories = ["新分类", "空分类"] } } };
        var firstPath = Path.Combine(root, "自选配置位置", "工作配置.json");
        ConfigurationStore.Save(firstPath, configuration);
        var loaded = ConfigurationStore.Load(firstPath, out bool legacy);
        Assert(!legacy && loaded.Software.Single().DefaultInstallDirectory == "D:\\" && loaded.Categories.Contains("空分类") && loaded.Schemes.ContainsKey("常用"),
            "chosen single JSON file restores software, directories, empty categories and schemes");
        var schemePath = Path.Combine(root, "old-schemes.json"); SchemeManager.SaveSchemes(configuration.Schemes, schemePath);
        Assert(ConfigurationStore.Load(schemePath, out legacy).Schemes.ContainsKey("常用") && legacy,
            "old scheme files can be explicitly imported without adopting their path for saves");

        Application.EnableVisualStyles();
        using var main = new MainForm(); main.ShowInTaskbar = false; main.Opacity = 0; main.Show(); Application.DoEvents();
        T Field<T>(object form, string name) => (T)form.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
        object? Invoke(object form, string name, params object?[] arguments) => form.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(form, arguments);
        var manager = Field<SoftwareManager>(main, "_softwareManager"); var handler = Field<SchemeHandler>(main, "_schemeHandler");
        Assert(manager.ActiveSoftwareItems.Count == 0 && manager.ActiveCategories.Count == 0 && !handler.GetSchemeNames().Any(),
            "startup ignores stale categories and even malformed old default data");
        Invoke(main, "LoadConfiguration", firstPath);
        Assert(manager.ActiveSoftwareItems.Count == 1 && manager.ActiveCategories.Contains("空分类") && handler.SchemeExists("常用"),
            "manual loading restores the whole workspace in the main window");
        var secondPath = Path.Combine(root, "第二份配置.json");
        ConfigurationStore.Save(secondPath, new() { Categories = ["第二份独有分类"] });
        Invoke(main, "LoadConfiguration", secondPath);
        Assert(manager.ActiveSoftwareItems.Count == 0 && manager.ActiveCategories.SequenceEqual(["第二份独有分类"]) && !handler.GetSchemeNames().Any(),
            "switching configurations replaces categories and schemes instead of merging previous data");
        var invalidPath = Path.Combine(root, "invalid.json"); File.WriteAllText(invalidPath, "{bad");
        bool rejected = false;
        try { Invoke(main, "LoadConfiguration", invalidPath); }
        catch (TargetInvocationException ex) when (ex.InnerException is JsonException) { rejected = true; }
        Assert(rejected && manager.ActiveCategories.SequenceEqual(["第二份独有分类"]), "failed loading leaves current configuration intact");
        Invoke(main, "LoadConfiguration", firstPath);
        manager.AddCategory("手动保存的新分类");
        Assert((bool)Invoke(main, "SaveConfiguration", false)! && ConfigurationStore.Load(firstPath, out _).Categories.Contains("手动保存的新分类") &&
            ConfigurationStore.Load(secondPath, out _).Categories.SequenceEqual(["第二份独有分类"]), "save writes only the selected configuration file and leaves other configurations untouched");
        using (var bitmap = new Bitmap(main.Width, main.Height)) { main.DrawToBitmap(bitmap, new Rectangle(Point.Empty, main.Size)); bitmap.Save(Path.Combine(root, "preview-configuration.png")); }
        using var fresh = new MainForm();
        Assert(Field<SoftwareManager>(fresh, "_softwareManager").ActiveCategories.Count == 0 &&
            File.ReadAllText(UserDataPaths.ReadPath("categories.json")) == "[\"旧分类\"]" &&
            File.ReadAllText(UserDataPaths.ReadPath("schemes.json")) == "invalid old JSON", "new sessions start blank and preserve legacy files for manual import");

        using var wizard = new BundleBuildForm(manager.ActiveSoftwareItems, handler, manager.ActiveCategories);
        wizard.ShowInTaskbar = false; wizard.Opacity = 0; wizard.Show(); Application.DoEvents();
        var checks = Field<ListView>(wizard, "checks");
        Assert(checks.Items.Count == 0 && Field<BundleBuildRequest?>(wizard, "prepared") == null, "new build wizard starts with no old check results");
        var request = (BundleBuildRequest)Invoke(wizard, "Prepare")!;
        Assert(request.Software.Count == 1 && request.Schemes["全部软件"].Count == 1, "packing all software works even when no items are checked");
        void Check()
        {
            var task = (Task)Invoke(wizard, "CheckAsync")!;
            var timer = Stopwatch.StartNew();
            while (!task.IsCompleted && timer.Elapsed < TimeSpan.FromSeconds(10)) { Application.DoEvents(); Thread.Sleep(5); }
            if (!task.IsCompleted) throw new Exception("Configuration check did not finish.");
            task.GetAwaiter().GetResult();
        }
        Invoke(wizard, "SetPage", 1); Check();
        Field<Button>(wizard, "next").PerformClick(); Application.DoEvents();
        Assert(Field<TabControl>(wizard, "pages").SelectedIndex == 2, "fresh checks allow the next button to open the generate page");
        Invoke(wizard, "SetPage", 0);
        Assert(checks.Items.Count == 0 && !Field<bool>(wizard, "checkedReady") && Field<BundleBuildRequest?>(wizard, "prepared") == null,
            "returning to content selection clears previous checks and prepared build data");
        Field<CheckBox>(wizard, "onlyChecked").Checked = true; Invoke(wizard, "SetPage", 1); Check();
        Assert(!Field<bool>(wizard, "checkedReady") && Field<Label>(wizard, "status").Text.Contains("未勾选软件"),
            "empty checked-only selections give a current actionable message instead of stale check information");
    }

    private static async Task Reject(Func<Task> action, string message)
    {
        bool rejected = false;
        try { await action(); }
        catch (Exception ex) when (ex is InvalidDataException or OperationCanceledException or IOException) { rejected = true; }
        Assert(rejected, message);
    }

    private static SoftwareItem Item(string name, int exitCode = 0, int order = 0, string? touch = null)
    {
        var file = Path.Combine(fixture, "Setup.exe");
        return new() { Id = Guid.NewGuid().ToString("N"), Name = name, Version = "1.0", Size = "", Description = "test", Category = "测试",
            FilePath = file, ResourceDirectory = fixture, SilentInstallArgs = $"--fake {exitCode} \"{Path.Combine(root, "ledger.txt")}\" {name} \"{touch ?? "-"}\" ignored 20",
            VerifiedSha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(file))), InstallOrder = order, AllowInteractiveFallback = false };
    }

    private static BundleBuildRequest Request(string output, params SoftwareItem[] items) => new()
    {
        Name = "测试离线合集", RunnerPath = runner, OutputPath = output, Software = items.ToList(), DefaultScheme = "全部",
        Categories = new() { "办公", "开发", "空分类" },
        Schemes = new() { ["全部"] = items.Select(s => s.Id).ToList(), ["基础"] = new() { items[0].Id } }
    };

    private static async Task RunAsync()
    {
        Assert(File.Exists(runner), "published self-contained runner exists");
        fixture = Path.Combine(root, "source files"); Directory.CreateDirectory(fixture);
        foreach (var file in Directory.EnumerateFiles(AppContext.BaseDirectory).Where(p => !p.EndsWith(".pdb")))
            File.Copy(file, Path.Combine(fixture, Path.GetFileName(file)), true);
        File.Copy(Environment.ProcessPath!, Path.Combine(fixture, "Setup.exe"), true);
        Directory.CreateDirectory(Path.Combine(fixture, "data"));
        await File.WriteAllTextAsync(Path.Combine(fixture, "data", "配置.txt"), "关联资源与中文路径测试");
        var first = Item("first", order: 50);
        var second = Item("second", order: 1); second.DependencyIds.Add(first.Id);
        first.Category = "办公"; second.Category = "开发";
        first.DefaultInstallDirectory = Path.Combine(root, "custom install 软件目录");
        first.DirectoryMode = InstallDirectoryMode.Custom; first.DirectoryParameter = "--target-dir \"{InstallDir}\"";
        first.DetectionFile = "{InstallDir}\\installed.marker";
        var output = Path.Combine(root, "软件合集.exe");
        await BundleBuilder.BuildAsync(Request(output, first, second));
        Assert(File.Exists(output), "bundle generated with Unicode filename");
        string extracted = Path.Combine(root, "extracted");
        using (var package = await BundleReader.OpenAsync(output))
        {
            Assert(package.Manifest.Schemes.Count == 2 && package.Manifest.Software.Count == 2, "multiple schemes round-trip");
            Assert(package.Manifest.Categories.Contains("空分类") && package.Manifest.Software.Select(s => s.Category).SequenceEqual(["办公", "开发"]), "empty and populated categories survive export");
            Assert(package.Manifest.Software[0].DefaultInstallDirectory == first.DefaultInstallDirectory && package.Manifest.Software[0].DirectoryMode == InstallDirectoryMode.Custom,
                "default install directory and argument mode survive export");
            Assert(package.Manifest.Software.All(s => !s.Installer.Contains(':')), "manifest contains internal paths");
            await package.ExtractAsync(package.Manifest.Software, extracted);
            Assert(File.ReadAllText(Path.Combine(extracted, first.Id, "data", "配置.txt")) == "关联资源与中文路径测试", "associated directory layout preserved");
            Assert(InstallationEngine.IncludeDependencies(package.Manifest.Software, [second.Id]).Count == 2, "dependency closure included");
            var outcomes = await InstallationEngine.RunAsync(package.Manifest.Software, extracted, Path.Combine(root, "logs-success"));
            Assert(outcomes.Count == 2 && outcomes.All(r => r.State == InstallState.Succeeded), "safe fixture installers execute successfully");
            Assert(File.ReadAllLines(Path.Combine(root, "ledger.txt")).SequenceEqual(["first", "second"]), "dependency order overrides numeric order");
            Assert(File.Exists(Path.Combine(root, "logs-success", "results.json")), "installation report persisted");
            Assert(File.Exists(Path.Combine(first.DefaultInstallDirectory, "installed.marker")), "installer actually receives configured directory with spaces and Unicode");
            var received = JsonSerializer.Deserialize<string[]>(File.ReadAllText(Path.Combine(first.DefaultInstallDirectory, "received.json")))!;
            Assert(received[Array.IndexOf(received, "--target-dir") + 1] == first.DefaultInstallDirectory, "directory argument reaches installer as one intact argument");
        }
        using (var input = File.OpenRead(output))
        {
            input.Position = input.Length - BundleReader.FooterSize; var footer = new byte[BundleReader.FooterSize]; input.ReadExactly(footer);
            long length = BitConverter.ToInt64(footer); input.Position = input.Length - BundleReader.FooterSize - length;
            using var memory = new MemoryStream(); await input.CopyToAsync(memory); memory.SetLength(length); memory.Position = 0;
            using var zip = new ZipArchive(memory, ZipArchiveMode.Read);
            int count = zip.Entries.Count(e => e.FullName.StartsWith("content/"));
            Assert(count == Directory.EnumerateFiles(fixture, "*", SearchOption.AllDirectories).Select(p => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(p)))).Distinct().Count(), "identical resources stored once");
        }
        // A generated single EXE must start by itself, with no adjacent DLLs or runtimeconfig.
        var report = Path.Combine(root, "runner-verification.json");
        using (var process = Process.Start(new ProcessStartInfo(output, $"--verify-package \"{report}\"") { UseShellExecute = false, CreateNoWindow = true })!)
        {
            using var limit = new CancellationTokenSource(TimeSpan.FromSeconds(90));
            await process.WaitForExitAsync(limit.Token);
            Assert(process.ExitCode == 0 && File.ReadAllText(report).Contains("\"success\":true"), "generated self-contained EXE starts and validates/extracts own payload");
        }
        var corrupt = Path.Combine(root, "corrupt.exe"); File.Copy(output, corrupt);
        using (var file = new FileStream(corrupt, FileMode.Open, FileAccess.ReadWrite)) { file.Position = file.Length - BundleReader.FooterSize - 10; int value = file.ReadByte(); file.Position--; file.WriteByte((byte)(value ^ 1)); }
        await Reject(async () => { using var _ = await BundleReader.OpenAsync(corrupt); }, "modified payload rejected");
        var oldHash = SHA256.HashData(File.ReadAllBytes(output));
        using (var cancelled = new CancellationTokenSource())
        {
            cancelled.Cancel();
            await Reject(() => BundleBuilder.BuildAsync(Request(output, first), token: cancelled.Token), "cancelled build rejected");
        }
        Assert(oldHash.SequenceEqual(SHA256.HashData(File.ReadAllBytes(output))), "cancelled build preserves previous output");
        var changed = first.Clone(); changed.FilePath = Path.Combine(root, "missing.exe");
        await Reject(() => BundleBuilder.BuildAsync(Request(output, changed)), "missing source file blocks build");
        Assert(oldHash.SequenceEqual(SHA256.HashData(File.ReadAllBytes(output))), "failed build preserves previous output");
        var unverified = first.Clone(); unverified.VerifiedSha256 = null;
        Assert(BundleBuilder.CheckItem(unverified).Count == 0, "unverified EXE is allowed without a manual verification step");
        unverified.RequiresNetwork = true;
        Assert(BundleBuilder.CheckItem(unverified).Count == 0, "network installer can be bundled with a network requirement");
        foreach (var bad in new[] { "../escape.exe", "/root.exe", "C:/escape.exe", "data/../a.exe", "data\\a.exe", "data/file:stream", "data/./a" })
            await Reject(() => { PackagePaths.Resolve(root, bad); return Task.CompletedTask; }, "reject path " + bad);
        var cyclic = new BundleSoftware { Id = "a", DependencyIds = ["b"] };
        await Reject(() => { InstallationEngine.Sort([cyclic, new() { Id = "b", DependencyIds = ["a"] }]); return Task.CompletedTask; }, "dependency cycles rejected");
        await FailureAndDetectionAsync();
        await InteractiveFallbackAsync();
        await CancelAndTimeoutAsync();
        await MalformedManifestAsync();
        await SizeLimitAsync();
        var legacy = "{\"Name\":\"Legacy\",\"Version\":\"1\",\"Size\":\"\",\"Description\":\"\",\"Category\":\"\",\"FilePath\":\"D:/old.exe\"}";
        Assert(JsonSerializer.Deserialize<SoftwareItem>(legacy)!.Id == JsonSerializer.Deserialize<SoftwareItem>(legacy)!.Id, "legacy scheme IDs migrate consistently");
        Assert(!Directory.EnumerateFiles(root, ".sia-*", SearchOption.TopDirectoryOnly).Any(), "temporary package files cleaned after failures");
    }

    private static async Task FailureAndDetectionAsync()
    {
        var failed = Item("failed", 17);
        var blocked = Item("blocked");
        var reboot = Item("reboot", 3010);
        blocked.DependencyIds.Add(failed.Id);
        var detectedFile = Path.Combine(root, "detected.txt");
        var detect = Item("detect", touch: detectedFile); detect.DetectionFile = detectedFile;
        var bundlePath = Path.Combine(root, "failure-tests.exe");
        await BundleBuilder.BuildAsync(Request(bundlePath, failed, blocked, reboot, detect));
        using var package = await BundleReader.OpenAsync(bundlePath);
        var work = Path.Combine(root, "fail-work"); await package.ExtractAsync(package.Manifest.Software, work);
        var outcomes = await InstallationEngine.RunAsync(package.Manifest.Software, work, Path.Combine(root, "logs-fail"));
        Assert(outcomes.Single(r => r.Id == failed.Id).State == InstallState.Failed, "non-success exit code marked failed");
        Assert(outcomes.Single(r => r.Id == blocked.Id).Message.Contains("前置"), "dependent installer skipped after prerequisite failure");
        Assert(outcomes.Single(r => r.Id == reboot.Id).State == InstallState.RebootRequired, "3010 recognized as reboot-required success");
        Assert(outcomes.Single(r => r.Id == detect.Id).State == InstallState.Succeeded, "post-install file detection passes");
        var again = await InstallationEngine.RunAsync([package.Manifest.Software.Single(s => s.Id == detect.Id)], work, Path.Combine(root, "logs-skip"));
        Assert(again.Single().State == InstallState.Skipped, "already installed software skipped");
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        var none = await InstallationEngine.RunAsync(package.Manifest.Software, work, Path.Combine(root, "logs-cancel"), token: cancelled.Token);
        Assert(none.Count == 0, "cancel prevents starting subsequent installers");
        var missing = package.Manifest.Software.Single(s => s.Id == detect.Id); missing.DetectionFile = Path.Combine(root, "missing.txt");
        var uncertain = await InstallationEngine.RunAsync([missing, package.Manifest.Software.Single(s => s.Id == reboot.Id)], work, Path.Combine(root, "logs-uncertain"));
        Assert(uncertain.Count == 1 && uncertain[0].State == InstallState.Uncertain, "failed post-install detection pauses queue");
    }

    private static void TestCategoryAndDirectorySettings()
    {
        TestAutomaticProfiles();
        var categoryPath = Path.Combine(root, "categories.json");
        var store = new CategoryStore(categoryPath);
        store.Save(["办公", "空分类", "办公"]);
        Assert(new CategoryStore(categoryPath).Load().SequenceEqual(["办公", "空分类"]), "category catalog persists empty categories across restart");
        var legacyPath = Path.Combine(root, "legacy-schemes.json");
        var item = new SoftwareItem { Name = "legacy", Version = "1", Size = "", Description = "", Category = "办公" };
        File.WriteAllText(legacyPath, JsonSerializer.Serialize(new Dictionary<string, List<SoftwareItem>> { ["旧方案"] = [item] }));
        var handler = new SchemeHandler(legacyPath);
        Assert(handler.GetCategories("旧方案").Contains("办公"), "legacy scheme categories migrate from software metadata");
        handler.SaveScheme("旧方案", [item], ["办公", "空分类"]);
        var loaded = new SchemeHandler(legacyPath);
        Assert(loaded.GetCategories("旧方案").Contains("空分类"), "saved schemes retain empty categories in versioned storage");
        loaded.GetCategories("旧方案").Clear();
        Assert(loaded.GetCategories("旧方案").Count == 2, "editing returned category list does not mutate saved scheme");
        loaded.SaveScheme("仅分类", [], ["空分类"]);
        Assert(new SchemeHandler(legacyPath).GetCategories("仅分类").Contains("空分类"), "scheme with no software still preserves category definitions");
        Assert(InstallDirectoryOptions.Normalize("D:") == "D:\\", "drive-letter entry D: normalizes to drive root");
        Assert(InstallDirectoryOptions.BuildArguments("/VERYSILENT", @"D:\Apps\My App", InstallDirectoryMode.InnoSetup, null, false) == "/VERYSILENT /DIR=\"D:\\Apps\\My App\"", "Inno directory argument quotes spaces correctly");
        Assert(InstallDirectoryOptions.BuildArguments("/S /NCRC", @"D:\Apps\My App", InstallDirectoryMode.Nsis, null, false) == @"/S /NCRC /D=D:\Apps\My App", "NSIS directory is final and unquoted");
        Assert(InstallDirectoryOptions.BuildArguments("", @"D:\Apps\My App", InstallDirectoryMode.MsiProperty, "APPLICATIONFOLDER", true) == "APPLICATIONFOLDER=\"D:\\Apps\\My App\"", "MSI directory uses configurable vendor property");
        foreach (var invalid in new[] { "D:relative", "relative", "D:\\Apps\\..\\Windows", "D:\\bad\"path" })
        {
            bool rejected = false;
            try { InstallDirectoryOptions.Validate(invalid, InstallDirectoryMode.InnoSetup, null, false); } catch (InvalidDataException) { rejected = true; }
            Assert(rejected, "invalid install directory rejected: " + invalid);
        }
        bool missingTemplate = false;
        try { InstallDirectoryOptions.BuildArguments("/S", "D:\\Apps", InstallDirectoryMode.Custom, "--dir=fixed", false); }
        catch (InvalidDataException) { missingTemplate = true; }
        Assert(missingTemplate, "custom directory template must reference actual setting");
        bool duplicateDirectory = false;
        try { InstallDirectoryOptions.BuildArguments("/S /D=C:\\Old", "D:\\Apps", InstallDirectoryMode.Nsis, null, false); }
        catch (InvalidDataException) { duplicateDirectory = true; }
        Assert(duplicateDirectory, "conflicting directory arguments rejected before installation");
        var missingDrive = "ZYXWVUTSRQPONMLKJIHGFED".Select(c => c + ":\\").FirstOrDefault(d => !Directory.Exists(d));
        if (missingDrive != null)
        {
            bool rejected = false;
            try { InstallDirectoryOptions.CheckTargetDrive(missingDrive); } catch (IOException) { rejected = true; }
            Assert(rejected, "missing target drive rejected without silently falling back to C");
        }
        Assert(InstallDirectoryOptions.ExpandDetectionPath(@"{InstallDir}\App.exe", @"D:\Apps\My App") == @"D:\Apps\My App\App.exe", "detection path follows configured install directory");
    }

    private static async Task MalformedManifestAsync()
    {
        var item = Item("bad");
        var manifest = new BundleManifest { Software = [new() { Id = item.Id, Installer = "../escape.exe", SuccessCodes = [0] }], Schemes = new() { ["当前方案"] = [item.Id] } };
        var path = Path.Combine(root, "malicious.exe");
        using var payload = new MemoryStream();
        using (var zip = new ZipArchive(payload, ZipArchiveMode.Create, true))
        { await using var entry = zip.CreateEntry("manifest.json").Open(); await JsonSerializer.SerializeAsync(entry, manifest); }
        using (var output = File.Create(path))
        { await output.WriteAsync(payload.ToArray()); await output.WriteAsync(BitConverter.GetBytes(payload.Length)); await output.WriteAsync(SHA256.HashData(payload.ToArray())); await output.WriteAsync(BundleReader.Magic); }
        await Reject(async () => { using var _ = await BundleReader.OpenAsync(path); }, "malicious manifest rejected even with valid payload digest");
    }

    private static void TestAutomaticProfiles()
    {
        var nsis = Path.Combine(root, "nsis.exe"); var inno = Path.Combine(root, "inno.exe"); var unknown = Path.Combine(root, "unknown.exe");
        File.WriteAllBytes(nsis, System.Text.Encoding.ASCII.GetBytes("MZ-test-NullsoftInst-test"));
        File.WriteAllBytes(inno, System.Text.Encoding.ASCII.GetBytes("MZ-test-Inno Setup Setup Data (6)-test"));
        File.WriteAllBytes(unknown, [77, 90, 0, 1, 2, 3]);
        SoftwareItem Recipe(string path) => new() { Name = "auto", Category = "测试", Version = "", Description = "", Size = "", FilePath = path };
        var nsisItem = Recipe(nsis); nsisItem.DefaultInstallDirectory = @"D:\Apps\My App";
        var nsisDescription = BundleBuilder.Describe(nsisItem);
        Assert(nsisDescription.Arguments == "/S" && nsisDescription.DirectoryMode == InstallDirectoryMode.Nsis && !nsisDescription.StartInteractively,
            "NSIS silent and directory arguments are detected without user input");
        var innoDescription = BundleBuilder.Describe(Recipe(inno));
        Assert(innoDescription.Arguments.Contains("/VERYSILENT") && innoDescription.Arguments.Contains("/NORESTART"), "Inno profile supplies common quiet arguments automatically");
        var unknownItem = Recipe(unknown); unknownItem.DefaultInstallDirectory = "D:";
        var unknownDescription = BundleBuilder.Describe(unknownItem);
        Assert(unknownDescription.StartInteractively && BundleBuilder.CheckItem(unknownItem).Count == 0, "unknown installer saves with no arguments and uses interactive installation");
        Assert(!InstallationEngine.CreateStartInfo(unknownDescription, unknown, root, true).Arguments.Contains("D:"), "unsupported directory is left for the interactive installer to choose");
        var msi = Path.Combine(root, "directory-test.msi");
        uint database = 0;
        try
        {
            uint error = TestMsiOpenDatabase(msi, new IntPtr(3), out database);
            if (error != 0) throw new Exception("Cannot create MSI fixture: " + error);
            TestSql(database, "CREATE TABLE `Directory` (`Directory` CHAR(72) NOT NULL, `Directory_Parent` CHAR(72), `DefaultDir` CHAR(255) NOT NULL LOCALIZABLE PRIMARY KEY `Directory`)");
            TestSql(database, "INSERT INTO `Directory` (`Directory`, `DefaultDir`) VALUES ('INSTALLFOLDER', 'Example')");
            if (TestMsiDatabaseCommit(database) != 0) throw new Exception("Cannot commit MSI fixture.");
        }
        finally { if (database != 0) TestMsiCloseHandle(database); }
        var msiItem = Recipe(msi); msiItem.DefaultInstallDirectory = @"D:\Apps\MSI App";
        var msiDescription = BundleBuilder.Describe(msiItem);
        Assert(msiDescription.IsMsi && !msiDescription.StartInteractively && msiDescription.DirectoryParameter == "INSTALLFOLDER",
            "MSI directory property discovered by read-only database inspection");
        var quietInfo = InstallationEngine.CreateStartInfo(msiDescription, msi, root, false);
        Assert(quietInfo.Arguments.Contains("/qn") && quietInfo.Arguments.Contains("INSTALLFOLDER=\"D:\\Apps\\MSI App\""), "MSI installation needs no user-written quiet or directory arguments");
        Assert(!InstallationEngine.CreateStartInfo(msiDescription, msi, root, true).Arguments.Contains("/qn"), "MSI interactive fallback omits silent options");
    }

    private static void TestSql(uint database, string sql)
    {
        uint error = TestMsiDatabaseOpenView(database, sql, out uint view);
        if (error != 0) throw new Exception("MSI SQL failed to open: " + error);
        try { error = TestMsiViewExecute(view, 0); if (error != 0) throw new Exception("MSI SQL failed to execute: " + error); }
        finally { TestMsiCloseHandle(view); }
    }

    private static async Task InteractiveFallbackAsync()
    {
        var source = Path.Combine(root, "fallback-source"); Directory.CreateDirectory(source);
        foreach (var file in Directory.EnumerateFiles(fixture)) File.Copy(file, Path.Combine(source, Path.GetFileName(file)));
        var filePath = Path.Combine(source, "FallbackFixture.exe"); File.Copy(Environment.ProcessPath!, filePath);
        var ledger = Path.Combine(root, "fallback-ledger.txt"); File.WriteAllText(Path.Combine(source, "scenario.txt"), ledger);
        var item = Item("fallback"); item.FilePath = filePath; item.ResourceDirectory = source; item.SilentInstallArgs = "--reject"; item.VerifiedSha256 = null; item.AllowInteractiveFallback = true;
        var bundle = Path.Combine(root, "fallback-test.exe"); await BundleBuilder.BuildAsync(Request(bundle, item));
        using var package = await BundleReader.OpenAsync(bundle);
        var work = Path.Combine(root, "fallback-work"); await package.ExtractAsync(package.Manifest.Software, work);
        var result = await InstallationEngine.RunAsync(package.Manifest.Software, work, Path.Combine(root, "fallback-logs"));
        Assert(result.Single().State == InstallState.Succeeded && File.ReadAllLines(ledger).SequenceEqual(["silent", "interactive"]), "invalid quiet parameters retry once as ordinary interactive installation");
        Assert(result.Single().Message.Contains("普通安装"), "fallback outcome clearly reports ordinary installation");
        File.Delete(ledger); var recipe = package.Manifest.Software.Single(); recipe.Arguments = "--cancel";
        var cancelled = await InstallationEngine.RunAsync([recipe], work, Path.Combine(root, "fallback-cancel-logs"));
        Assert(cancelled.Single().State == InstallState.Failed && File.ReadAllLines(ledger).SequenceEqual(["silent"]), "user cancellation does not reopen installer");
        File.Delete(ledger); recipe.Arguments = ""; recipe.StartInteractively = true;
        var ordinary = await InstallationEngine.RunAsync([recipe], work, Path.Combine(root, "ordinary-logs"));
        Assert(ordinary.Single().State == InstallState.Succeeded && File.ReadAllLines(ledger).SequenceEqual(["interactive"]), "installer without quiet arguments runs once in ordinary mode");
    }

    [DllImport("msi.dll", EntryPoint = "MsiOpenDatabaseW", CharSet = CharSet.Unicode)]
    private static extern uint TestMsiOpenDatabase(string path, IntPtr mode, out uint database);
    [DllImport("msi.dll", EntryPoint = "MsiDatabaseOpenViewW", CharSet = CharSet.Unicode)]
    private static extern uint TestMsiDatabaseOpenView(uint database, string query, out uint view);
    [DllImport("msi.dll", EntryPoint = "MsiViewExecute")] private static extern uint TestMsiViewExecute(uint view, uint record);
    [DllImport("msi.dll", EntryPoint = "MsiDatabaseCommit")] private static extern uint TestMsiDatabaseCommit(uint database);
    [DllImport("msi.dll", EntryPoint = "MsiCloseHandle")] private static extern uint TestMsiCloseHandle(uint handle);

    private static async Task CancelAndTimeoutAsync()
    {
        var slow = Item("slow");
        slow.SilentInstallArgs = slow.SilentInstallArgs![..^2] + "500";
        var next = Item("afterSlow");
        var path = Path.Combine(root, "timeout-tests.exe");
        await BundleBuilder.BuildAsync(Request(path, slow, next));
        using var package = await BundleReader.OpenAsync(path);
        var work = Path.Combine(root, "timeout-work"); await package.ExtractAsync(package.Manifest.Software, work);
        using (var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(150)))
        {
            var cancelled = await InstallationEngine.RunAsync(package.Manifest.Software, work, Path.Combine(root, "logs-cancel-current"), token: cancellation.Token);
            Assert(cancelled.Count == 1 && cancelled[0].State == InstallState.Succeeded, "cancel lets current installer finish and stops next item");
        }
        var timed = package.Manifest.Software[0];
        timed.TimeoutMinutes = 1; timed.Arguments = timed.Arguments[..^3] + "65000";
        var watch = Stopwatch.StartNew();
        Console.WriteLine("Checking real one-minute installer timeout (no third-party installs)...");
        var uncertain = await InstallationEngine.RunAsync(package.Manifest.Software, work, Path.Combine(root, "logs-timeout"));
        Assert(uncertain.Count == 1 && uncertain[0].State == InstallState.Uncertain && watch.Elapsed >= TimeSpan.FromSeconds(59), "real timeout pauses queue without force-killing installer");
        await Task.Delay(TimeSpan.FromSeconds(6)); // The harmless fixture exits by itself; leave no background process.
    }

    private static async Task SizeLimitAsync()
    {
        var stub = Path.Combine(root, "huge-runner.exe");
        using (var file = File.Create(stub)) { file.Write([77, 90]); file.SetLength(BundleBuilder.MaxBundleBytes); }
        var request = Request(Path.Combine(root, "oversize.exe"), Item("oversize"));
        await Reject(() => BundleBuilder.BuildAsync(new() { Name = request.Name, Software = request.Software, Schemes = request.Schemes,
            DefaultScheme = request.DefaultScheme, RunnerPath = stub, OutputPath = request.OutputPath }), "oversized single EXE blocked before output replacement");
        File.Delete(stub);
    }

    private static void RenderPreviews()
    {
        Application.EnableVisualStyles();
        // Surface initialization errors in the console instead of a blocking error dialog.
        _ = new SoftwareManager(); _ = new SchemeHandler();
        using var main = new MainForm();
        var manager = (SoftwareManager)typeof(MainForm).GetField("_softwareManager", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(main)!;
        manager.ActiveCategories.Add("空分类预览");
        var rebuild = typeof(MainForm).GetMethod("RebuildTreeView", BindingFlags.Instance | BindingFlags.NonPublic)!;
        rebuild.Invoke(main, null); rebuild.Invoke(main, null);
        var tree = (TreeView)typeof(MainForm).GetField("categoryTreeView", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(main)!;
        Assert(tree.Nodes[0].Nodes.Cast<TreeNode>().Any(n => n.Text == "空分类预览"), "empty category survives repeated main-list rebuilds");
        using var editor = new SoftwareEditForm();
        using var wizard = new BundleBuildForm([], new SchemeHandler());
        foreach (var (form, name) in new (Form, string)[] { (main, "main"), (editor, "editor"), (wizard, "wizard") })
        {
            form.ShowInTaskbar = false; form.Opacity = 0; form.Show(); Application.DoEvents(); form.PerformLayout();
            using var bitmap = new Bitmap(form.Width, form.Height); form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.Size));
            bitmap.Save(Path.Combine(root, "preview-" + name + ".png"));
        }
        var editLayout = (TableLayoutPanel)typeof(SoftwareEditForm).GetField("layout", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(editor)!;
        Assert(editLayout.RowCount == 10 && !editLayout.AutoScroll, "editor contains only the ten requested rows without a scrolling form");
        Assert(editLayout.RowStyles.Cast<RowStyle>().All(r => r.SizeType == SizeType.Absolute) && editor.ClientSize.Height <= 550,
            "compact editor uses content-sized fixed rows rather than stretched spacing");
        Assert(editLayout.Controls.OfType<TextBox>().All(t => t.Width <= 310), "editor text fields use approximately half the old width");
        editor.Size = editor.MinimumSize; Application.DoEvents(); editor.PerformLayout();
        Assert(editLayout.Controls.Cast<Control>().Where(c => c.Visible).All(c => editLayout.ClientRectangle.Contains(c.Bounds)), "all editor rows remain inside the minimum window size");
        using (var bitmap = new Bitmap(editor.Width, editor.Height)) { editor.DrawToBitmap(bitmap, new Rectangle(Point.Empty, editor.Size)); bitmap.Save(Path.Combine(root, "preview-editor-minimum.png")); }
        var customItem = new SoftwareItem { Name = "目录参数布局", Version = "", Size = "", Description = "", Category = "未分类" };
        customItem.DirectoryMode = InstallDirectoryMode.Custom; customItem.DirectoryParameter = "--dir={InstallDir}";
        using var customEditor = new SoftwareEditForm(customItem);
        customEditor.ShowInTaskbar = false; customEditor.Opacity = 0; customEditor.Show(); Application.DoEvents();
        var customLayout = (TableLayoutPanel)typeof(SoftwareEditForm).GetField("layout", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(customEditor)!;
        Assert(customLayout.RowStyles[5].Height == 92 && customEditor.ClientSize.Height == editor.ClientSize.Height + 24,
            "saved custom directory mode expands its parameter row before the editor is shown");
        using (var bitmap = new Bitmap(customEditor.Width, customEditor.Height)) { customEditor.DrawToBitmap(bitmap, new Rectangle(Point.Empty, customEditor.Size)); bitmap.Save(Path.Combine(root, "preview-editor-custom.png")); }
        var mode = (ComboBox)typeof(SoftwareEditForm).GetField("directoryMode", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(customEditor)!;
        mode.SelectedIndex = 0; Application.DoEvents();
        Assert(customEditor.ClientSize.Height == editor.ClientSize.Height, "returning to automatic directory mode removes the extra parameter row space");
        Assert(true, "WinForms dialogs construct and render without exceptions");
    }

    private static void RenderRunnerAndWizard()
    {
        using var form = new InstallForm(Path.Combine(root, "软件合集.exe"));
        form.ShowInTaskbar = false; form.Opacity = 0; form.Show();
        var watch = Stopwatch.StartNew();
        var install = (Button)typeof(InstallForm).GetField("install", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
        while (!install.Enabled && watch.Elapsed < TimeSpan.FromSeconds(30)) { Application.DoEvents(); Thread.Sleep(10); }
        Assert(install.Enabled, "runner UI loads package and enables installation without executing installers");
        var list = (ListView)typeof(InstallForm).GetField("list", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
        Assert(list.Items.Count == 2 && list.CheckedItems.Count == 2, "runner displays all apps and preselects default scheme");
        var categories = (ComboBox)typeof(InstallForm).GetField("categories", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
        var ids = (HashSet<string>)typeof(InstallForm).GetField("selectedIds", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
        categories.SelectedIndex = categories.Items.Cast<object>().ToList().FindIndex(c => c.ToString() == "分类：空分类");
        Assert(list.Items.Count == 0 && ids.Count == 2, "empty category visible in runner and filtering preserves hidden selections");
        categories.SelectedIndex = categories.Items.Cast<object>().ToList().FindIndex(c => c.ToString() == "分类：办公");
        Assert(list.Items.Count == 1 && list.Items[0].SubItems[4].Text == "办公", "runner filters software by persisted category");
        list.Items[0].Checked = false;
        categories.SelectedIndex = 0;
        Assert(list.Items.Count == 2 && list.CheckedItems.Count == 1 && ids.Count == 1, "changing category does not lose checkbox edits");
        var schemes = (ComboBox)typeof(InstallForm).GetField("schemes", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
        schemes.SelectedItem = "基础";
        Assert(list.CheckedItems.Count == 1 && list.CheckedItems[0].SubItems[4].Text == "办公", "scheme selection works after category filtering");
        schemes.SelectedItem = "全部";
        using (var bitmap = new Bitmap(form.Width, form.Height)) { form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.Size)); bitmap.Save(Path.Combine(root, "preview-runner.png")); }
        var current = new[] { Item("first"), Item("second") }; foreach (var item in current) item.IsSelected = true;
        using var wizard = new BundleBuildForm(current, new SchemeHandler(), ["办公", "开发", "空分类"]);
        wizard.ShowInTaskbar = false; wizard.Opacity = 0; wizard.Show();
        var plan = (BundleBuildRequest)typeof(BundleBuildForm).GetMethod("Prepare", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(wizard, null)!;
        Assert(plan.Software.Count == 2 && plan.Schemes["当前方案"].Count == 2, "wizard derives package content and default scheme from current selections");
        Assert(plan.Categories.Contains("空分类"), "generation wizard carries standalone category catalog");
        typeof(BundleBuildForm).GetMethod("SetPage", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(wizard, [2]); Application.DoEvents();
        using (var bitmap = new Bitmap(wizard.Width, wizard.Height)) { wizard.DrawToBitmap(bitmap, new Rectangle(Point.Empty, wizard.Size)); bitmap.Save(Path.Combine(root, "preview-generate.png")); }
        var edited = Item("editor");
        using var editor = new SoftwareEditForm(edited, current, ["办公", "空分类"]);
        editor.ShowInTaskbar = false; editor.Opacity = 0; editor.Show();
        T Field<T>(string name) => (T)typeof(SoftwareEditForm).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(editor)!;
        Field<TextBox>("installDirectory").Text = "D:";
        Field<ComboBox>("directoryMode").SelectedIndex = 1;
        Field<ComboBox>("category").Text = "空分类";
        Assert(typeof(SoftwareEditForm).GetField("verified", BindingFlags.Instance | BindingFlags.NonPublic) == null, "simple editor has no manual verification checkbox");
        Assert(Field<Label>("directoryHint").Text.Contains("/DIR"), "manual directory types include explanatory text");
        using (var bitmap = new Bitmap(editor.Width, editor.Height)) { editor.DrawToBitmap(bitmap, new Rectangle(Point.Empty, editor.Size)); bitmap.Save(Path.Combine(root, "preview-directory-editor.png")); }
        Field<Button>("okButton").PerformClick(); Application.DoEvents();
        Assert(edited.DefaultInstallDirectory == "D:\\" && edited.DirectoryMode == InstallDirectoryMode.InnoSetup && edited.Category == "空分类" && edited.VerifiedSha256 == null,
            "editor saves directory, category and invalidated verification correctly");
    }
}
