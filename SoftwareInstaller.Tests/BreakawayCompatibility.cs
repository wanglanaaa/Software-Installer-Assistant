using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using SoftwareInstaller.Packaging;

namespace SoftwareInstaller.Tests;

internal static partial class Program
{
    private static int RunBreakawayFixture(string marker, int milliseconds)
    {
        File.WriteAllText(marker + ".parent", Environment.ProcessId.ToString());
        var startup = new FixtureStartup { Size = Marshal.SizeOf<FixtureStartup>() };
        var command = new StringBuilder($"\"{Environment.ProcessPath}\" --process-fixture sleep {milliseconds} \"{marker}\"");
        if (!FixtureCreateProcess(Environment.ProcessPath!, command, IntPtr.Zero, IntPtr.Zero, false,
            0x01000000 | 0x08000000, IntPtr.Zero, AppContext.BaseDirectory, ref startup, out var child))
        { File.WriteAllText(marker + ".error", Marshal.GetLastWin32Error().ToString()); return 9; }
        FixtureCloseHandle(child.Thread); FixtureCloseHandle(child.Process);
        Thread.Sleep(650); // Allow monitoring to observe the launcher before it exits.
        return 0;
    }

    private static async Task TestBreakawayCompatibilityAsync()
    {
        var blocked = Path.Combine(root, "blocked.pid");
        var strict = new BundleSoftware { Id = "strict", Name = "受限进程替身", Installer = "fixture.exe", AllowInteractiveFallback = false,
            Arguments = $"--breakaway-fixture \"{blocked}\" 100" };
        var results = await InstallationEngine.RunAsync([strict], root, Path.Combine(root, "strict-logs"),
            resolveInstaller: _ => Environment.ProcessPath!, cancelCurrent: true);
        Assert(results.Single().State == InstallState.Failed && results[0].ExitCode == 9 && File.ReadAllText(blocked + ".error") == "5",
            "a restrictive job reproduces access denied when a wrapper starts a breakaway child, without treating exit code 9 as success");

        var source = Path.Combine(root, "wrapper"); Directory.CreateDirectory(source);
        foreach (var file in Directory.EnumerateFiles(AppContext.BaseDirectory)) File.Copy(file, Path.Combine(source, Path.GetFileName(file)));
        var executable = Path.Combine(source, "InteractiveFixture.exe"); File.Copy(Environment.ProcessPath!, executable);
        var manual = new BundleSoftware { Id = "wrapper", Name = "绿色单文件替身", Installer = "InteractiveFixture.exe", StartInteractively = true };
        var nextMarker = Path.Combine(root, "next.pid");
        var next = new BundleSoftware { Id = "next", Name = "下一项", Installer = "fixture.exe",
            Arguments = $"--process-fixture sleep 20 \"{nextMarker}\"" };
        string scenarioPath = Path.Combine(source, "breakaway-scenario.txt");
        string marker = Path.Combine(root, "compatible.pid");
        File.WriteAllLines(scenarioPath, [marker, "1800"]);
        var clock = Stopwatch.StartNew();
        results = await InstallationEngine.RunAsync([manual, next], root, Path.Combine(root, "compatible-logs"),
            resolveInstaller: item => item == manual ? executable : Environment.ProcessPath!, cancelCurrent: true).WaitAsync(TimeSpan.FromSeconds(8));
        Assert(results.Count == 2 && results.All(r => r.State == InstallState.Succeeded) && File.Exists(nextMarker)
            && !File.Exists(marker + ".error") && clock.ElapsedMilliseconds >= 1500,
            "an interactive wrapper can launch its breakaway child and the queue waits for that child before continuing");
        Assert(!File.ReadAllText(Path.Combine(root, "compatible-logs", "session.log")).Contains("正在等待安装子进程：PID=" + File.ReadAllText(marker + ".parent") + " "),
            "an exited launcher is never reported as its own remaining child");

        marker = Path.Combine(root, "cancelled.pid"); File.WriteAllLines(scenarioPath, [marker, "30000"]);
        using var cancellation = new CancellationTokenSource();
        var task = InstallationEngine.RunAsync([manual, next], root, Path.Combine(root, "cancelled-logs"), token: cancellation.Token,
            resolveInstaller: item => item == manual ? executable : Environment.ProcessPath!, cancelCurrent: true);
        Process? child = null;
        try
        {
            clock.Restart();
            while (!File.Exists(marker) && clock.Elapsed < TimeSpan.FromSeconds(8)) await Task.Delay(20);
            child = Process.GetProcessById(int.Parse(File.ReadAllText(marker))); _ = child.SafeHandle;
            using var parent = Process.GetProcessById(int.Parse(File.ReadAllText(marker + ".parent"))); _ = parent.SafeHandle;
            await parent.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(8));
            Assert(!task.IsCompleted && !child.HasExited, "monitoring retains a breakaway child after the wrapper exits");
            cancellation.Cancel(); results = await task.WaitAsync(TimeSpan.FromSeconds(8));
            Assert(results.Count == 1 && results[0].State == InstallState.Cancelled && child.HasExited,
                "cancellation also terminates the observed breakaway child and stops the remaining queue");
        }
        finally
        {
            cancellation.Cancel();
            if (child != null) { if (!child.HasExited) { child.Kill(); await child.WaitForExitAsync(); } child.Dispose(); }
        }
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct FixtureStartup
    {
        public int Size;
        public IntPtr Reserved, Desktop, Title;
        public uint X, Y, XSize, YSize, XCountChars, YCountChars, FillAttribute, Flags;
        public ushort ShowWindow, ReservedSize;
        public IntPtr ReservedBytes, StandardInput, StandardOutput, StandardError;
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct FixtureProcess { public IntPtr Process, Thread; public uint ProcessId, ThreadId; }
    [DllImport("kernel32.dll", EntryPoint = "CreateProcessW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool FixtureCreateProcess(string application, StringBuilder command, IntPtr processAttributes,
        IntPtr threadAttributes, bool inheritHandles, uint flags, IntPtr environment, string directory,
        ref FixtureStartup startup, out FixtureProcess process);
    [DllImport("kernel32.dll", EntryPoint = "CloseHandle")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool FixtureCloseHandle(IntPtr handle);
}
