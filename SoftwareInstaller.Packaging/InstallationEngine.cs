using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using SoftwareInstaller.Models;

namespace SoftwareInstaller.Packaging;

public enum InstallState { Succeeded, Skipped, Failed, RebootRequired, Uncertain }
public sealed record InstallResult(string Id, string Name, InstallState State, string Message, int? ExitCode = null);
public sealed record InstallUpdate(string Id, string Message, int Completed, int Total, InstallResult? Result = null);

public static class InstallationEngine
{
    public static List<BundleSoftware> Sort(IEnumerable<BundleSoftware> software)
    {
        var items = software.ToList(); var byId = items.ToDictionary(s => s.Id);
        var states = new Dictionary<string, int>(); var sorted = new List<BundleSoftware>();
        void Visit(BundleSoftware item)
        {
            if (states.TryGetValue(item.Id, out var state)) { if (state == 1) throw new InvalidDataException("安装依赖存在循环。"); return; }
            states[item.Id] = 1;
            foreach (var dependency in item.DependencyIds)
            {
                if (!byId.TryGetValue(dependency, out var parent)) throw new InvalidDataException($"{item.Name} 缺少前置软件，请将其一起打包。");
                Visit(parent);
            }
            states[item.Id] = 2; sorted.Add(item);
        }
        foreach (var item in items.OrderBy(s => s.Order)) Visit(item);
        return sorted;
    }

    public static List<BundleSoftware> IncludeDependencies(IEnumerable<BundleSoftware> all, IEnumerable<string> selectedIds)
    {
        var items = Sort(all); var selected = selectedIds.ToHashSet();
        if (selected.Any(id => !items.Any(s => s.Id == id))) throw new InvalidDataException("选择了未知的软件。");
        void Add(BundleSoftware item) { foreach (var id in item.DependencyIds) if (selected.Add(id)) Add(items.Single(s => s.Id == id)); }
        foreach (var item in items.Where(s => selected.Contains(s.Id)).ToList()) Add(item);
        return items.Where(s => selected.Contains(s.Id)).ToList();
    }

    public static bool HasDetection(BundleSoftware item) => !string.IsNullOrWhiteSpace(item.DetectionFile) || !string.IsNullOrWhiteSpace(item.MsiProductCode);
    public static bool IsInstalled(BundleSoftware item)
    {
        if (!HasDetection(item)) return false;
        if (!string.IsNullOrWhiteSpace(item.MsiProductCode) && MsiQueryProductState("{" + Guid.Parse(item.MsiProductCode).ToString().ToUpperInvariant() + "}") != 5) return false;
        if (!string.IsNullOrWhiteSpace(item.DetectionFile))
        {
            var path = InstallDirectoryOptions.ExpandDetectionPath(item.DetectionFile, item.DefaultInstallDirectory);
            if (!File.Exists(path)) return false;
            if (!string.IsNullOrWhiteSpace(item.DetectionVersion))
            {
                var version = FileVersionInfo.GetVersionInfo(path).FileVersion?.Replace(',', '.').Replace(" ", "");
                if (!Version.TryParse(version, out var actual) || !Version.TryParse(item.DetectionVersion, out var expected) || actual < expected) return false;
            }
        }
        return true;
    }

    public static ProcessStartInfo CreateStartInfo(BundleSoftware item, string installer, string logDirectory, bool interactive,
        bool discardDirectory = false, bool requestElevation = false)
    {
        string directory = discardDirectory ? "" : item.DefaultInstallDirectory ?? "";
        var mode = discardDirectory ? InstallDirectoryMode.Automatic : item.DirectoryMode;
        string arguments = InstallDirectoryOptions.BuildArguments(interactive ? "" : item.Arguments, directory, mode, item.DirectoryParameter, item.IsMsi);
        var info = new ProcessStartInfo
        {
            FileName = item.IsMsi ? Path.Combine(Environment.SystemDirectory, "msiexec.exe") : installer,
            WorkingDirectory = Path.GetDirectoryName(Path.GetFullPath(installer))!,
            UseShellExecute = requestElevation, CreateNoWindow = !interactive, WindowStyle = ProcessWindowStyle.Normal
        };
        if (requestElevation) info.Verb = "runas";
        info.Arguments = item.IsMsi ? $"/i {InstallDirectoryOptions.Quote(installer)} {arguments} {(interactive ? "" : "/qn")} /norestart REBOOT=ReallySuppress /L*v {InstallDirectoryOptions.Quote(Path.Combine(logDirectory, item.Id + (interactive ? ".interactive.msi.log" : ".msi.log")))}" : arguments;
        return info;
    }

    public static async Task<List<InstallResult>> RunAsync(IEnumerable<BundleSoftware> software, string workDirectory, string logDirectory,
        IProgress<InstallUpdate>? progress = null, CancellationToken token = default,
        Func<BundleSoftware, string>? resolveInstaller = null, bool requestElevation = false)
    {
        var items = Sort(software); Directory.CreateDirectory(logDirectory); var results = new List<InstallResult>();
        await File.WriteAllTextAsync(Path.Combine(logDirectory, "session.log"), $"开始：{DateTimeOffset.Now:O}{Environment.NewLine}");
        foreach (var item in items)
        {
            if (token.IsCancellationRequested) break;
            progress?.Report(new(item.Id, $"正在检查：{item.Name}", results.Count, items.Count));
            InstallResult result;
            if (item.DependencyIds.Any(id => !results.Any(r => r.Id == id && r.State is InstallState.Succeeded or InstallState.Skipped or InstallState.RebootRequired)))
                result = new(item.Id, item.Name, InstallState.Failed, "前置软件安装失败，已跳过。");
            else result = await InstallOneAsync(item, results.Count, items.Count);
            results.Add(result);
            await LogAsync($"{result.Name}: {result.State} {result.Message}");
            await File.WriteAllTextAsync(Path.Combine(logDirectory, "results.json"), JsonSerializer.Serialize(results, BundleBuilder.JsonOptions));
            progress?.Report(new(item.Id, result.Message, results.Count, items.Count, result));
            if (result.State == InstallState.Uncertain || result.ExitCode == 1641) break;
        }
        return results;

        Task LogAsync(string message) => File.AppendAllTextAsync(Path.Combine(logDirectory, "session.log"), $"{DateTimeOffset.Now:O} {message}{Environment.NewLine}");

        async Task<InstallResult> InstallOneAsync(BundleSoftware item, int completed, int total)
        {
            bool started = false;
            try
            {
                if (IsInstalled(item)) return new(item.Id, item.Name, InstallState.Skipped, "已满足安装检测规则，跳过。");
                var installer = resolveInstaller?.Invoke(item) ?? PackagePaths.Resolve(PackagePaths.Resolve(workDirectory, item.Id), item.Installer);
                if (!File.Exists(installer)) throw new FileNotFoundException("安装文件不存在。", installer);
                bool directoryNeedsUi = !string.IsNullOrWhiteSpace(item.DefaultInstallDirectory) && item.DirectoryMode == InstallDirectoryMode.Automatic;
                if (!directoryNeedsUi) InstallDirectoryOptions.CheckTargetDrive(item.DefaultInstallDirectory);
                bool interactive = item.StartInteractively || (!item.IsMsi && string.IsNullOrWhiteSpace(item.Arguments)) || directoryNeedsUi;
                bool fallback = false;
                while (true)
                {
                    var message = interactive ? $"请在 {item.Name} 的安装界面中完成操作" : $"正在自动安装：{item.Name}";
                    if (fallback) message = $"{item.Name} 的静默参数未成功，已切换普通安装，请按窗口提示操作";
                    if (directoryNeedsUi || fallback && !string.IsNullOrWhiteSpace(item.DefaultInstallDirectory)) message += $"；请确认安装位置 {item.DefaultInstallDirectory}";
                    progress?.Report(new(item.Id, message, completed, total));
                    await LogAsync(message);
                    var info = CreateStartInfo(item, installer, logDirectory, interactive, fallback, requestElevation);
                    using var process = Process.Start(info) ?? throw new IOException("无法启动安装程序。");
                    started = true;
                    using var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(item.TimeoutMinutes));
                    try { await process.WaitForExitAsync(timeout.Token); }
                    catch (OperationCanceledException) { return new(item.Id, item.Name, InstallState.Uncertain, "等待超时，安装器可能仍在运行。队列已暂停，请检查安装窗口。"); }
                    int code = process.ExitCode;
                    await LogAsync($"{item.Name} {(interactive ? "普通安装" : "静默安装")} 返回码 {code}");
                    bool reboot = item.RebootCodes.Contains(code);
                    bool success = item.SuccessCodes.Contains(code) || reboot;
                    if (!success && !interactive && item.AllowInteractiveFallback && code is not (1602 or 1223) && !token.IsCancellationRequested)
                    {
                        interactive = true; fallback = true;
                        continue;
                    }
                    if (!success) return new(item.Id, item.Name, InstallState.Failed, code is 1602 or 1223 ? "安装已取消。" : $"安装器返回失败代码 {code}。", code);
                    if (reboot) return new(item.Id, item.Name, InstallState.RebootRequired, "安装器报告完成，需要重启后复核。", code);
                    if (HasDetection(item))
                    {
                        bool installed = IsInstalled(item);
                        for (int retry = 0; retry < 10 && !installed; retry++) { await Task.Delay(1000); installed = IsInstalled(item); }
                        if (!installed) return new(item.Id, item.Name, InstallState.Uncertain, "安装器已退出但检测未通过，队列已暂停，请检查软件是否安装完成。", code);
                    }
                    string done = fallback ? "已切换普通安装并完成。" : interactive ? "普通安装已完成。" : "自动安装已完成。";
                    return new(item.Id, item.Name, InstallState.Succeeded, done + (HasDetection(item) ? "检测通过。" : "安装器报告成功。"), code);
                }
            }
            catch (Exception ex) { return new(item.Id, item.Name, started ? InstallState.Uncertain : InstallState.Failed, started ? "安装开始后出现异常，队列已暂停：" + ex.Message : ex.Message); }
        }
    }

    [DllImport("msi.dll", EntryPoint = "MsiQueryProductStateW", CharSet = CharSet.Unicode)]
    private static extern int MsiQueryProductState(string product);
}
