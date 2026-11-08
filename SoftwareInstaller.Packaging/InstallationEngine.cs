using System.Diagnostics;
using System.Runtime.InteropServices;
using SoftwareInstaller.Models;

namespace SoftwareInstaller.Packaging;

public enum InstallState { Succeeded, Skipped, Failed, RebootRequired, Uncertain, Cancelled }
public sealed record InstallResult(string Id, string Name, InstallState State, string Message, int? ExitCode = null);
public enum InstallPhase { Checking, Installing, WaitingForUser, Verifying, Preparing }
public sealed record InstallUpdate(string Id, string Message, int Completed, int Total, InstallResult? Result = null, InstallPhase Phase = InstallPhase.Checking, bool CanTerminate = false)
{
    public string StateText => Result != null ? Result.State switch
    {
        InstallState.Succeeded => "已完成", InstallState.Skipped => "已跳过", InstallState.Failed => "失败",
        InstallState.RebootRequired => "需要重启", InstallState.Cancelled => "已取消", _ => "待人工检查"
    } : Phase switch
    {
        InstallPhase.Installing => "正在安装", InstallPhase.WaitingForUser => "等待你操作", InstallPhase.Verifying => "验证安装结果", InstallPhase.Preparing => "准备资源", _ => "检查安装条件"
    };
}

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
        => InstallationPlan.Create(all).Select(selectedIds).Items.ToList();

    public static bool HasDetection(BundleSoftware item) => !string.IsNullOrWhiteSpace(item.DetectionFile) || !string.IsNullOrWhiteSpace(item.MsiProductCode);
    public static bool IsInstalled(BundleSoftware item)
    {
        if (!HasDetection(item)) return false;
        if (!string.IsNullOrWhiteSpace(item.MsiProductCode) && MsiQueryProductState("{" + Guid.Parse(item.MsiProductCode).ToString().ToUpperInvariant() + "}") != 5) return false;
        if (!string.IsNullOrWhiteSpace(item.DetectionFile))
        {
            var path = InstallDirectoryOptions.ExpandDetectionPath(item.DetectionFile!, item.DefaultInstallDirectory);
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

    public static Task<List<InstallResult>> RunAsync(IEnumerable<BundleSoftware> software, string workDirectory, string logDirectory,
        IProgress<InstallUpdate>? progress = null, CancellationToken token = default,
        Func<BundleSoftware, string>? resolveInstaller = null, bool requestElevation = false, InstallationControl? control = null,
        Func<BundleSoftware, CancellationToken, Task>? prepareInstaller = null, bool cancelCurrent = false)
        => RunAsync(InstallationPlan.Create(software), workDirectory, logDirectory, progress, token, resolveInstaller, requestElevation, control, prepareInstaller, cancelCurrent);

    public static async Task<List<InstallResult>> RunAsync(InstallationPlan plan, string workDirectory, string logDirectory,
        IProgress<InstallUpdate>? progress = null, CancellationToken token = default,
        Func<BundleSoftware, string>? resolveInstaller = null, bool requestElevation = false, InstallationControl? control = null,
        Func<BundleSoftware, CancellationToken, Task>? prepareInstaller = null, bool cancelCurrent = false)
    {
        var items = plan.Items; Directory.CreateDirectory(logDirectory); var results = new List<InstallResult>();
        var resultsById = new Dictionary<string, InstallResult>(StringComparer.Ordinal);
        await RunnerRuntime.WriteTextAsync(Path.Combine(logDirectory, "session.log"), $"开始：{DateTimeOffset.Now:O}{Environment.NewLine}");
        foreach (var item in items)
        {
            if (token.IsCancellationRequested || control?.Token.IsCancellationRequested == true) break;
            progress?.Report(new(item.Id, $"正在检查：{item.Name}", results.Count, items.Count));
            InstallResult result;
            if (item.DependencyIds.Any(id => !resultsById.TryGetValue(id, out var dependency) || dependency.State is not (InstallState.Succeeded or InstallState.Skipped or InstallState.RebootRequired)))
                result = new(item.Id, item.Name, InstallState.Failed, "前置软件安装失败，已跳过。");
            else result = await InstallOneAsync(item, results.Count, items.Count);
            results.Add(result);
            resultsById.Add(result.Id, result);
            await LogAsync($"{result.Name}: {result.State} {result.Message}");
            await RunnerRuntime.WriteTextAsync(Path.Combine(logDirectory, "results.json"), RunnerRuntime.Json(results));
            progress?.Report(new(item.Id, result.Message, results.Count, items.Count, result));
            if (result.State is InstallState.Uncertain or InstallState.Cancelled || result.ExitCode == 1641) break;
        }
        return results;

        Task LogAsync(string message) => RunnerRuntime.WriteTextAsync(Path.Combine(logDirectory, "session.log"), $"{DateTimeOffset.Now:O} {message}{Environment.NewLine}", append: true);

        async Task<InstallResult> InstallOneAsync(BundleSoftware item, int completed, int total)
        {
            bool started = false;
            try
            {
                if (IsInstalled(item)) return new(item.Id, item.Name, InstallState.Skipped, "已满足安装检测规则，跳过。");
                if (prepareInstaller != null)
                {
                    using var preparing = CancellationTokenSource.CreateLinkedTokenSource(token, control?.Token ?? CancellationToken.None);
                    preparing.Token.ThrowIfCancellationRequested();
                    progress?.Report(new(item.Id, $"正在准备：{item.Name}", completed, total, Phase: InstallPhase.Preparing));
                    await prepareInstaller(item, preparing.Token);
                    preparing.Token.ThrowIfCancellationRequested();
                }
                var installer = resolveInstaller?.Invoke(item) ?? PackagePaths.Resolve(PackagePaths.Resolve(workDirectory, item.Id), item.Installer);
                if (!File.Exists(installer)) throw new FileNotFoundException("安装文件不存在。", installer);
                bool directoryNeedsUi = !string.IsNullOrWhiteSpace(item.DefaultInstallDirectory) && item.DirectoryMode == InstallDirectoryMode.Automatic;
                if (!directoryNeedsUi) InstallDirectoryOptions.CheckTargetDrive(item.DefaultInstallDirectory);
                bool interactive = item.StartInteractively || (!item.IsMsi && string.IsNullOrWhiteSpace(item.Arguments)) || directoryNeedsUi;
                bool fallback = false;
                while (true)
                {
                    if (token.IsCancellationRequested || control?.Token.IsCancellationRequested == true)
                        return new(item.Id, item.Name, InstallState.Cancelled, "已取消，未启动安装器。");
                    var message = interactive ? $"请在 {item.Name} 的安装界面中完成操作" : $"正在自动安装：{item.Name}";
                    if (fallback) message = $"{item.Name} 的静默参数未成功，已切换普通安装，请按窗口提示操作";
                    if (directoryNeedsUi || fallback && !string.IsNullOrWhiteSpace(item.DefaultInstallDirectory)) message += $"；请确认安装位置 {item.DefaultInstallDirectory}";
                    progress?.Report(new(item.Id, message, completed, total, Phase: interactive ? InstallPhase.WaitingForUser : InstallPhase.Installing));
                    await LogAsync(message);
                    var info = CreateStartInfo(item, installer, logDirectory, interactive, fallback, requestElevation);
                    if (token.IsCancellationRequested || control?.Token.IsCancellationRequested == true)
                        return new(item.Id, item.Name, InstallState.Cancelled, "已取消，未启动安装器。");
                    using var session = InstallerProcess.Start(info, trackChildren: !requestElevation && (!item.IsMsi || cancelCurrent), killOnClose: cancelCurrent);
                    var process = session.Process;
                    started = true;
                    await LogAsync($"{item.Name} PID={process.Id}；安装器={installer}；子进程管理={session.CanTerminate}；资源目录={workDirectory}");
                    if (session.TrackingWarning != null) await LogAsync(session.TrackingWarning);
                    progress?.Report(new(item.Id, message, completed, total, Phase: interactive ? InstallPhase.WaitingForUser : InstallPhase.Installing, CanTerminate: session.CanTerminate));
                    using var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(item.TimeoutMinutes));
                    using var waiting = CancellationTokenSource.CreateLinkedTokenSource(timeout.Token, control?.Token ?? CancellationToken.None, cancelCurrent ? token : CancellationToken.None);
                    try { await session.WaitAsync(waiting.Token); }
                    catch (OperationCanceledException)
                    {
                        string stopped = "等待超时，安装器可能仍在运行。";
                        if (control?.TerminationRequested == true || cancelCurrent && (token.IsCancellationRequested || control?.Token.IsCancellationRequested == true || timeout.IsCancellationRequested))
                        {
                            try
                            {
                                await session.TerminateAsync();
                                stopped = timeout.IsCancellationRequested ? "等待超时，已终止当前安装器及其受管理的子进程。" : "已取消当前安装器及其受管理的子进程。";
                                if (cancelCurrent && !item.IsMsi && !timeout.IsCancellationRequested)
                                    return new(item.Id, item.Name, InstallState.Cancelled, stopped + " 后续任务已停止。");
                                if (item.IsMsi) stopped += " Windows Installer 服务可能仍在回滚，请等待其完成。";
                            }
                            catch (Exception ex) { stopped = "未能确认安装器全部终止：" + ex.Message; }
                        }
                        else if (control?.Token.IsCancellationRequested == true) stopped = "已停止等待，当前安装器可能继续运行。";
                        return new(item.Id, item.Name, InstallState.Uncertain, stopped + " 队列已停止，资源暂时保留。");
                    }
                    if ((cancelCurrent && token.IsCancellationRequested) || control?.Token.IsCancellationRequested == true)
                        return cancelCurrent
                            ? new(item.Id, item.Name, InstallState.Cancelled, "已取消，受管理的安装进程已退出，后续任务已停止。")
                            : new(item.Id, item.Name, InstallState.Uncertain, "已停止等待，安装结果待人工检查，资源暂时保留。");
                    if (session.TrackingWarning != null)
                        return new(item.Id, item.Name, InstallState.Uncertain, session.TrackingWarning + "；主进程已退出，子进程状态待确认，资源暂时保留。", process.ExitCode);
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
                        progress?.Report(new(item.Id, $"正在验证 {item.Name} 的安装结果", completed, total, Phase: InstallPhase.Verifying));
                        bool installed = IsInstalled(item);
                        using var verifying = CancellationTokenSource.CreateLinkedTokenSource(token, control?.Token ?? CancellationToken.None);
                        for (int retry = 0; retry < 10 && !installed; retry++) { await Task.Delay(1000, verifying.Token); installed = IsInstalled(item); }
                        if (!installed) return new(item.Id, item.Name, InstallState.Uncertain, "安装器已退出但检测未通过，队列已暂停，请检查软件是否安装完成。", code);
                    }
                    string done = fallback ? "已切换普通安装并完成。" : interactive ? "普通安装已完成。" : "自动安装已完成。";
                    return new(item.Id, item.Name, InstallState.Succeeded, done + (HasDetection(item) ? "检测通过。" : "安装器报告成功。"), code);
                }
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested || control?.Token.IsCancellationRequested == true)
            { return new(item.Id, item.Name, started ? InstallState.Uncertain : InstallState.Cancelled, started ? "已停止等待，安装结果待人工检查。" : "已取消准备资源，未启动安装器。"); }
            catch (Exception ex) { return new(item.Id, item.Name, started ? InstallState.Uncertain : InstallState.Failed, started ? "安装开始后出现异常，队列已暂停：" + ex.Message : ex.Message); }
        }
    }

    [DllImport("msi.dll", EntryPoint = "MsiQueryProductStateW", CharSet = CharSet.Unicode)]
    private static extern int MsiQueryProductState(string product);
}
