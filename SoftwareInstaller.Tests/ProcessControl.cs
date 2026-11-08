using System.Diagnostics;
using SoftwareInstaller.Packaging;

namespace SoftwareInstaller.Tests;

internal static partial class Program
{
    private static int RunProcessFixture(string[] args)
    {
        if (args[1] == "sleep")
        {
            File.WriteAllText(args[3], Environment.ProcessId.ToString());
            Thread.Sleep(int.Parse(args[2]));
            return 0;
        }
        File.WriteAllText(args[3] + ".parent", Environment.ProcessId.ToString());
        using var child = Process.Start(new ProcessStartInfo(Environment.ProcessPath!, $"--process-fixture sleep {args[2]} \"{args[3]}\"") { UseShellExecute = false, CreateNoWindow = true })!;
        if (args[1] != "exit-parent") child.WaitForExit();
        return 0;
    }

    private sealed class InlineProgress(Action<InstallUpdate> action) : IProgress<InstallUpdate>
    {
        public void Report(InstallUpdate value) => action(value);
    }

    private static async Task TestProcessControlAsync()
    {
        BundleSoftware Item(string mode, int milliseconds, string marker) => new()
        {
            Id = Guid.NewGuid().ToString("N"), Name = mode, Installer = "setup.exe", AllowInteractiveFallback = true,
            Arguments = $"--process-fixture {mode} {milliseconds} \"{marker}\""
        };
        Task<List<InstallResult>> Run(BundleSoftware[] items, string name, InstallationControl? control = null,
            CancellationToken token = default, IProgress<InstallUpdate>? progress = null, bool cancelCurrent = false) => InstallationEngine.RunAsync(items, root,
                Path.Combine(root, name), progress, token, resolveInstaller: _ => Environment.ProcessPath!, control: control, cancelCurrent: cancelCurrent);
        async Task<int> ReadPid(string path)
        {
            var watch = Stopwatch.StartNew();
            while (watch.Elapsed < TimeSpan.FromSeconds(8))
            {
                if (File.Exists(path) && int.TryParse(File.ReadAllText(path), out var pid)) return pid;
                await Task.Delay(20);
            }
            throw new Exception("Fixture did not start: " + path);
        }
        async Task Cleanup(string marker)
        {
            foreach (var path in new[] { marker + ".parent", marker })
            {
                if (!File.Exists(path) || !int.TryParse(File.ReadAllText(path), out var pid)) continue;
                try
                {
                    using var process = Process.GetProcessById(pid);
                    if (!process.HasExited && string.Equals(process.MainModule?.FileName, Environment.ProcessPath, StringComparison.OrdinalIgnoreCase))
                    { process.Kill(); await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(5)); }
                }
                catch (ArgumentException) { }
            }
        }

        var marker = Path.Combine(root, "never-started.pid");
        using (var stop = new CancellationTokenSource())
        {
            var result = await Run([Item("sleep", 100, marker)], "logs-before-start", token: stop.Token,
                progress: new InlineProgress(p => { if (p.Phase == InstallPhase.Installing) stop.Cancel(); }));
            Assert(result.Single().State == InstallState.Cancelled && !File.Exists(marker), "stopping during preparation prevents the installer from starting");
        }

        marker = Path.Combine(root, "finish-current.pid");
        using (var stop = new CancellationTokenSource())
        {
            var task = Run([Item("sleep", 400, marker), Item("sleep", 100, marker + ".next")], "logs-stop-queue", token: stop.Token);
            await ReadPid(marker); stop.Cancel();
            var result = await task.WaitAsync(TimeSpan.FromSeconds(8));
            Assert(result.Count == 1 && result[0].State == InstallState.Succeeded && !File.Exists(marker + ".next"), "stopping the queue lets the current installer finish and skips the next");
        }

        marker = Path.Combine(root, "early-parent.pid");
        try
        {
            var task = Run([Item("exit-parent", 1000, marker)], "logs-early-parent");
            using var child = Process.GetProcessById(await ReadPid(marker));
            Assert(!task.IsCompleted, "an early launcher exit does not finish the installation while its child runs");
            var result = await task.WaitAsync(TimeSpan.FromSeconds(8));
            Assert(result.Single().State == InstallState.Succeeded && child.HasExited, "installation completes after all managed child processes exit");
        }
        finally { await Cleanup(marker); }

        marker = Path.Combine(root, "cancel-managed-tree.pid");
        using (var stop = new CancellationTokenSource())
        {
            try
            {
                var task = Run([Item("exit-parent", 30000, marker), Item("sleep", 100, marker + ".next")],
                    "logs-cancel-managed-tree", token: stop.Token, cancelCurrent: true);
                using var child = Process.GetProcessById(await ReadPid(marker));
                _ = child.SafeHandle;
                stop.Cancel();
                var result = await task.WaitAsync(TimeSpan.FromSeconds(8));
                Assert(result.Single().State == InstallState.Cancelled && child.HasExited && !File.Exists(marker + ".next"),
                    "runner cancellation kills children even after their launcher exits and never starts the next item");
            }
            finally { await Cleanup(marker); }
        }

        foreach (bool force in new[] { false, true })
        {
            marker = Path.Combine(root, force ? "terminate-tree.pid" : "detach-tree.pid");
            using var control = new InstallationControl();
            try
            {
                bool canTerminate = false;
                var task = Run([Item("tree", 30000, marker), Item("sleep", 100, marker + ".next")],
                    force ? "logs-terminate" : "logs-detach", control, progress: new InlineProgress(p => { if (p.CanTerminate) canTerminate = true; }));
                using var child = Process.GetProcessById(await ReadPid(marker));
                using var parent = Process.GetProcessById(await ReadPid(marker + ".parent"));
                _ = child.SafeHandle; _ = parent.SafeHandle;
                if (force) control.TerminateCurrent(); else control.StopWaiting();
                var result = await task.WaitAsync(TimeSpan.FromSeconds(8));
                Assert(canTerminate && result.Single().State == InstallState.Uncertain && !File.Exists(marker + ".next"),
                    force ? "forced termination pauses the queue without fallback or retry" : "stopping the wait returns promptly and pauses the queue");
                Assert(force ? parent.HasExited && child.HasExited : !parent.HasExited && !child.HasExited,
                    force ? "forced termination ends both the installer and its child" : "stopping the wait leaves the current installer and child running");
                var log = File.ReadAllText(Path.Combine(root, force ? "logs-terminate" : "logs-detach", "session.log"));
                Assert(log.Contains("PID=" + parent.Id) && log.Contains(Environment.ProcessPath!) && log.Contains(root), "interrupted installation logs include its PID, installer and resource directory");
            }
            finally { await Cleanup(marker); }
        }
    }
}
