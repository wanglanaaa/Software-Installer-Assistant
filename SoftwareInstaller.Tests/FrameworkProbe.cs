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
        try { CheckAsync(args).GetAwaiter().GetResult(); return 0; }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
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
