using System.Diagnostics;
using System.Reflection;
using SoftwareInstaller.Packaging;

namespace SoftwareInstaller.Tests;

internal static partial class Program
{
    private static async Task TestProcessTreeModeAsync()
    {
        var type = typeof(InstallationEngine).Assembly.GetType("SoftwareInstaller.Packaging.InstallerProcess")!;
        object Start(string mode, int milliseconds, string marker) => type.GetMethod("Start")!.Invoke(null,
            [new ProcessStartInfo(Environment.ProcessPath!, $"--process-fixture {mode} {milliseconds} \"{marker}\"")
                { UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = AppContext.BaseDirectory }, true, true, false, true])!;
        Task Wait(object session, CancellationToken token) => (Task)type.GetMethod("WaitAsync")!.Invoke(session, [token, null])!;
        Task Terminate(object session) => (Task)type.GetMethod("TerminateAsync")!.Invoke(session, null)!;
        async Task<int> ReadPid(string marker)
        {
            var clock = Stopwatch.StartNew();
            while (clock.Elapsed < TimeSpan.FromSeconds(8))
            {
                if (File.Exists(marker) && int.TryParse(File.ReadAllText(marker), out var pid)) return pid;
                await Task.Delay(20);
            }
            throw new Exception("Missing process tree marker: " + marker);
        }

        var marker = Path.Combine(root, "tree-only-finish.pid");
        var session = Start("exit-parent", 1200, marker);
        try
        {
            var wait = Wait(session, CancellationToken.None);
            using var child = Process.GetProcessById(await ReadPid(marker)); _ = child.SafeHandle;
            Assert((bool)type.GetProperty("CanTerminate")!.GetValue(session)! &&
                ((string)type.GetProperty("TrackingNotice")!.GetValue(session)!).Contains("兼容 Win7"),
                "Win7 tracking starts without requiring a job and reports that compatibility mode is active");
            Assert(!wait.IsCompleted, "without a job, an early launcher exit still waits for its installation child");
            await wait.WaitAsync(TimeSpan.FromSeconds(8));
            Assert(child.HasExited, "without a job, completion is reported after the child exits");
        }
        finally { await Terminate(session); ((IDisposable)session).Dispose(); }

        foreach (string mode in new[] { "tree", "exit-parent" })
        {
            marker = Path.Combine(root, "tree-only-cancel-" + mode + ".pid");
            session = Start(mode, 30000, marker);
            using var stop = new CancellationTokenSource();
            try
            {
                var wait = Wait(session, stop.Token);
                using var child = Process.GetProcessById(await ReadPid(marker)); _ = child.SafeHandle;
                var parent = (Process)type.GetProperty("Process")!.GetValue(session)!;
                if (mode == "exit-parent") await parent.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(8));
                stop.Cancel();
                try { await wait; } catch (OperationCanceledException) { }
                await Terminate(session).WaitAsync(TimeSpan.FromSeconds(8));
                Assert(parent.HasExited && child.HasExited, "without a job, cancellation ends the child and its " + mode + " launcher");
            }
            finally { await Terminate(session); ((IDisposable)session).Dispose(); }
        }
    }
}
