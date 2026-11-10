#if FRAMEWORK_PROBE
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using SoftwareInstaller.Packaging;

// Compile this small probe against the actual net48 runner to exercise its linked runtime.
internal static class FrameworkProbe
{
    private static int Main(string[] args)
    {
        AppDomain.CurrentDomain.AssemblyResolve += (_, name) =>
        {
            if (new AssemblyName(name.Name).Name != "Newtonsoft.Json") return null;
            using var source = typeof(BundleReader).Assembly.GetManifestResourceStream("SoftwareInstaller.Newtonsoft.Json.dll");
            using var bytes = new MemoryStream(); source.CopyTo(bytes); return Assembly.Load(bytes.ToArray());
        };
        try
        {
            if (args.Length == 3 && args[0] == "--installer-flow") CheckInstallerFlowAsync(args[1], args[2]).GetAwaiter().GetResult();
            else if (args.Length == 3 && args[0] == "--breakaway") CheckBreakawayAsync(args[1], args[2]).GetAwaiter().GetResult();
            else if (args.Length == 3 && args[0] == "--process-tree") CheckProcessTreeAsync(args[1], args[2]).GetAwaiter().GetResult();
            else CheckAsync(args).GetAwaiter().GetResult();
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }

    private static async Task CheckInstallerFlowAsync(string output, string fixture)
    {
        string root = Path.GetFullPath(output); Directory.CreateDirectory(root);
        fixture = Path.GetFullPath(fixture);
        var applicationDirectory = Path.Combine(root, "application"); Directory.CreateDirectory(applicationDirectory);
        foreach (var file in Directory.EnumerateFiles(Path.GetDirectoryName(fixture)))
            File.Copy(file, Path.Combine(applicationDirectory, Path.GetFileName(file)), true);
        var interactive = Path.Combine(applicationDirectory, "InteractiveFixture.exe"); File.Copy(fixture, interactive, true);
        var marker = Path.Combine(root, "manual.txt");
        File.WriteAllText(Path.Combine(applicationDirectory, "interactive-scenario.txt"), marker);
        var manual = new BundleSoftware { Id = "manual", Name = "手动程序", Installer = "setup.exe", StartInteractively = true,
            DefaultInstallDirectory = @"D:\", DetectionFile = @"{InstallDir}\old.marker" };
        var nextMarker = Path.Combine(root, "next.pid");
        var next = new BundleSoftware { Id = "next", Name = "下一项", Installer = "setup.exe",
            Arguments = "--process-fixture sleep 20 \"" + nextMarker + "\"" };
        var task = InstallationEngine.RunAsync(new[] { manual, next }, root, Path.Combine(root, "manual-logs"),
            resolveInstaller: item => item == manual ? interactive : fixture, cancelCurrent: true, useInstallerDefaults: true);
        if (await Task.WhenAny(task, Task.Delay(8000)) != task) throw new Exception("Framework interactive queue did not finish");
        var results = await task;
        if (results.Count != 2 || results.Any(r => r.State != InstallState.Succeeded) || !File.Exists(nextMarker) ||
            File.ReadAllText(marker) != "opened\nclosed") throw new Exception("Framework interactive completion did not continue the queue");
        Console.WriteLine("PASS: actual Framework runner continues after an interactive program exits using installer defaults");

        marker = Path.Combine(root, "installed.pid");
        File.WriteAllText(marker + ".application", Path.Combine(applicationDirectory, Path.GetFileName(fixture)));
        var launcher = new BundleSoftware { Id = "launcher", Name = "安装后启动程序", Installer = "setup.exe",
            Arguments = "--process-fixture launch-installed 30000 \"" + marker + "\"" };
        try
        {
            task = InstallationEngine.RunAsync(new[] { launcher, next }, root, Path.Combine(root, "installed-logs"),
                resolveInstaller: _ => fixture, cancelCurrent: true, useInstallerDefaults: true);
            if (await Task.WhenAny(task, Task.Delay(8000)) != task) throw new Exception("Framework waited on an installed application");
            results = await task;
            var clock = Stopwatch.StartNew();
            while (!File.Exists(marker) && clock.Elapsed < TimeSpan.FromSeconds(8)) await Task.Delay(20);
            using var application = Process.GetProcessById(int.Parse(File.ReadAllText(marker)));
            if (results.Count != 2 || results.Any(r => r.State != InstallState.Succeeded) || application.HasExited)
                throw new Exception("Framework blocked the queue or killed the installed application");
            Console.WriteLine("PASS: actual Framework runner leaves installed applications running without blocking the queue");
        }
        finally
        {
            if (File.Exists(marker))
            {
                try { using var application = Process.GetProcessById(int.Parse(File.ReadAllText(marker))); if (!application.HasExited) { application.Kill(); application.WaitForExit(); } }
                catch (ArgumentException) { }
            }
        }
    }

    private static async Task CheckBreakawayAsync(string output, string fixture)
    {
        string root = Path.GetFullPath(output); Directory.CreateDirectory(root);
        fixture = Path.GetFullPath(fixture);
        var source = Path.Combine(root, "wrapper"); Directory.CreateDirectory(source);
        foreach (var file in Directory.EnumerateFiles(Path.GetDirectoryName(fixture)))
            File.Copy(file, Path.Combine(source, Path.GetFileName(file)), true);
        var executable = Path.Combine(source, "InteractiveFixture.exe"); File.Copy(fixture, executable, true);
        var scenario = Path.Combine(source, "breakaway-scenario.txt");
        var item = new BundleSoftware { Id = "wrapper", Name = "Framework breakaway wrapper", Installer = "setup.exe", StartInteractively = true };
        var nextMarker = Path.Combine(root, "next.pid");
        var next = new BundleSoftware { Id = "next", Name = "next", Installer = "setup.exe",
            Arguments = "--process-fixture sleep 20 \"" + nextMarker + "\"" };
        var marker = Path.Combine(root, "finished.pid"); File.WriteAllLines(scenario, new[] { marker, "1800" });
        var clock = Stopwatch.StartNew();
        var task = InstallationEngine.RunAsync(new[] { item, next }, root, Path.Combine(root, "complete-logs"),
            resolveInstaller: selected => selected == item ? executable : fixture, cancelCurrent: true);
        if (await Task.WhenAny(task, Task.Delay(8000)) != task) throw new Exception("Framework breakaway wait timed out");
        var results = await task;
        if (results.Count != 2 || results.Any(r => r.State != InstallState.Succeeded) || !File.Exists(nextMarker) ||
            File.Exists(marker + ".error") || clock.ElapsedMilliseconds < 1500)
            throw new Exception("Framework failed to launch or wait for the breakaway child");
        Console.WriteLine("PASS: actual Framework runner launches and waits for the breakaway child before continuing");

        marker = Path.Combine(root, "cancelled.pid"); File.WriteAllLines(scenario, new[] { marker, "30000" });
        nextMarker = Path.Combine(root, "cancelled-next.pid"); next.Arguments = "--process-fixture sleep 20 \"" + nextMarker + "\"";
        using var cancellation = new CancellationTokenSource();
        Process child = null;
        try
        {
            task = InstallationEngine.RunAsync(new[] { item, next }, root, Path.Combine(root, "cancel-logs"), token: cancellation.Token,
                resolveInstaller: selected => selected == item ? executable : fixture, cancelCurrent: true);
            clock.Restart();
            while (!File.Exists(marker) && clock.Elapsed < TimeSpan.FromSeconds(8)) await Task.Delay(20);
            child = Process.GetProcessById(int.Parse(File.ReadAllText(marker))); _ = child.Handle;
            using var parent = Process.GetProcessById(int.Parse(File.ReadAllText(marker + ".parent"))); _ = parent.Handle;
            while (!parent.HasExited && clock.Elapsed < TimeSpan.FromSeconds(8)) await Task.Delay(20);
            if (!parent.HasExited || task.IsCompleted || child.HasExited) throw new Exception("Framework lost the child after its launcher exited");
            cancellation.Cancel();
            if (await Task.WhenAny(task, Task.Delay(8000)) != task) throw new Exception("Framework breakaway cancellation timed out");
            results = await task;
            if (results.Count != 1 || results[0].State != InstallState.Cancelled || !child.HasExited || File.Exists(nextMarker))
                throw new Exception("Framework cancellation left the breakaway child or continued the queue");
            Console.WriteLine("PASS: actual Framework runner cancels the breakaway child after the launcher exits and stops the queue");
        }
        finally
        {
            cancellation.Cancel();
            if (child != null) { if (!child.HasExited) { child.Kill(); child.WaitForExit(); } child.Dispose(); }
        }
    }

    private static async Task CheckProcessTreeAsync(string output, string fixture)
    {
        string root = Path.GetFullPath(output); Directory.CreateDirectory(root);
        fixture = Path.GetFullPath(fixture);
        var type = typeof(InstallationEngine).Assembly.GetType("SoftwareInstaller.Packaging.InstallerProcess");
        foreach (bool cancel in new[] { false, true })
        {
            string marker = Path.Combine(root, cancel ? "cancel.pid" : "finish.pid");
            var info = new ProcessStartInfo(fixture, "--process-fixture exit-parent " + (cancel ? "30000" : "1200") + " \"" + marker + "\"")
                { UseShellExecute = false, CreateNoWindow = true, WorkingDirectory = Path.GetDirectoryName(fixture) };
            var session = type.GetMethod("Start").Invoke(null, new object[] { info, true, true, false, true });
            using var stop = new CancellationTokenSource();
            try
            {
                var wait = (Task)type.GetMethod("WaitAsync").Invoke(session, new object[] { stop.Token, null });
                var clock = Stopwatch.StartNew();
                while (!File.Exists(marker) && clock.Elapsed < TimeSpan.FromSeconds(8)) await Task.Delay(20);
                using var child = Process.GetProcessById(int.Parse(File.ReadAllText(marker))); _ = child.Handle;
                var parent = (Process)type.GetProperty("Process").GetValue(session);
                while (!parent.HasExited && clock.Elapsed < TimeSpan.FromSeconds(8)) await Task.Delay(20);
                if (!parent.HasExited || wait.IsCompleted || child.HasExited || !(bool)type.GetProperty("CanTerminate").GetValue(session))
                    throw new Exception("Framework process-tree tracking lost the child or cannot terminate it");
                if (cancel)
                {
                    stop.Cancel();
                    try { await wait; } catch (OperationCanceledException) { }
                    var terminate = (Task)type.GetMethod("TerminateAsync").Invoke(session, null);
                    if (await Task.WhenAny(terminate, Task.Delay(8000)) != terminate) throw new Exception("Framework tree termination timed out");
                    await terminate;
                }
                else
                {
                    if (await Task.WhenAny(wait, Task.Delay(8000)) != wait) throw new Exception("Framework tree completion timed out");
                    await wait;
                }
                if (!child.HasExited) throw new Exception("Framework process-tree tracking left the child running");
                Console.WriteLine(cancel ? "PASS: actual Framework runner cancels the remaining child without a job" :
                    "PASS: actual Framework runner waits for the remaining child without a job");
            }
            finally
            {
                await (Task)type.GetMethod("TerminateAsync").Invoke(session, null);
                ((IDisposable)session).Dispose();
            }
        }
    }

    private static async Task CheckAsync(string[] args)
    {
        string root = Path.GetFullPath(args[0]); Directory.CreateDirectory(root);
        using var package = await BundleReader.OpenAsync(args[1]);
        var item = package.Manifest.Software[0];
        await package.ExtractAsync(new[] { item }, Path.Combine(root, "extracted"));
        if (Directory.GetDirectories(Path.Combine(root, "extracted")).Length != 1 ||
            File.ReadAllBytes(Path.Combine(root, "extracted", item.Id, "setup.exe")).Length != 5)
            throw new Exception("Unexpected demand extraction result");
        Console.WriteLine("PASS: actual Framework runner reads JSON and extracts/hashes only selected software");

        var marker = Path.Combine(root, "process.pid");
        var nextMarker = marker + ".next";
        var fixture = Path.GetFullPath(args[2]);
        var first = new BundleSoftware { Id = Guid.NewGuid().ToString("N"), Name = "Framework cancel", Installer = "setup.exe",
            Arguments = "--process-fixture exit-parent 30000 \"" + marker + "\"", AllowInteractiveFallback = true };
        var next = new BundleSoftware { Id = Guid.NewGuid().ToString("N"), Name = "Next", Installer = "setup.exe",
            Arguments = "--process-fixture sleep 10 \"" + nextMarker + "\"" };
        int childPid = 0;
        try
        {
            using var cancellation = new CancellationTokenSource();
            var task = InstallationEngine.RunAsync(new[] { first, next }, root, Path.Combine(root, "logs"),
                token: cancellation.Token, resolveInstaller: _ => fixture, cancelCurrent: true);
            var clock = Stopwatch.StartNew();
            while (!File.Exists(marker) && clock.Elapsed < TimeSpan.FromSeconds(10)) await Task.Delay(20);
            childPid = int.Parse(File.ReadAllText(marker));
            using var child = Process.GetProcessById(childPid);
            _ = child.Handle;
            cancellation.Cancel();
            if (await Task.WhenAny(task, Task.Delay(8000)) != task) throw new Exception("Cancellation did not finish promptly");
            var results = await task;
            if (results.Count != 1 || results[0].State != InstallState.Cancelled || !child.HasExited || File.Exists(nextMarker))
                throw new Exception("Framework cancellation left a process or continued the queue");
            Console.WriteLine("PASS: actual Framework runner cancels the remaining child and stops the next installer");
        }
        finally
        {
            if (childPid != 0)
            {
                try { using var child = Process.GetProcessById(childPid); if (!child.HasExited) child.Kill(); }
                catch (ArgumentException) { }
            }
        }
    }
}
#endif
