using System.Diagnostics;
using SoftwareInstaller.Models;
using SoftwareInstaller.Packaging;

namespace SoftwareInstaller.Tests;

internal static partial class Program
{
    private static async Task TestInstallerFlowAsync()
    {
        var fixtureDirectory = Path.Combine(root, "application"); Directory.CreateDirectory(fixtureDirectory);
        foreach (var file in Directory.EnumerateFiles(AppContext.BaseDirectory))
            File.Copy(file, Path.Combine(fixtureDirectory, Path.GetFileName(file)));
        var interactive = Path.Combine(fixtureDirectory, "InteractiveFixture.exe"); File.Copy(Environment.ProcessPath!, interactive);
        string marker = Path.Combine(root, "interactive.txt");
        File.WriteAllText(Path.Combine(fixtureDirectory, "interactive-scenario.txt"), marker);
        var manual = new BundleSoftware { Id = Guid.NewGuid().ToString("N"), Name = "手动程序", Installer = "InteractiveFixture.exe",
            StartInteractively = true, DefaultInstallDirectory = @"D:\", DirectoryMode = InstallDirectoryMode.Nsis,
            DetectionFile = @"{InstallDir}\old.marker" };
        var nextMarker = Path.Combine(root, "next.pid");
        var next = new BundleSoftware { Id = Guid.NewGuid().ToString("N"), Name = "下一项", Installer = "setup.exe",
            Arguments = $"--process-fixture sleep 20 \"{nextMarker}\"", DefaultInstallDirectory = @"?:\invalid",
            DirectoryMode = InstallDirectoryMode.Nsis };
        var updates = new List<InstallUpdate>();
        var configuration = new SoftwareItem { Name = "旧路径配置", Version = "", Category = "测试", Size = "", Description = "",
            FilePath = Environment.ProcessPath, SilentInstallArgs = "--quiet", DefaultInstallDirectory = @"?:\invalid", DirectoryMode = InstallDirectoryMode.Automatic };
        var bundleDescription = BundleBuilder.DescribeForBundle(configuration);
        Assert(BundleBuilder.CheckItem(configuration, useInstallerDefaults: true).Count == 0 && !bundleDescription.StartInteractively
            && string.IsNullOrEmpty(bundleDescription.DefaultInstallDirectory) && configuration.DefaultInstallDirectory == @"?:\invalid",
            "generation ignores obsolete directory settings without forcing a silent package into interactive mode or changing the saved configuration");
        var task = InstallationEngine.RunAsync(new[] { manual, next }, root, Path.Combine(root, "logs-interactive"),
            new InlineProgress(p => updates.Add(p)), resolveInstaller: item => item == manual ? interactive : Environment.ProcessPath!,
            cancelCurrent: true, useInstallerDefaults: true);
        var clock = Stopwatch.StartNew();
        while (!File.Exists(marker) && clock.Elapsed < TimeSpan.FromSeconds(8)) await Task.Delay(20);
        Assert(File.Exists(marker) && !task.IsCompleted && !File.Exists(nextMarker), "a standalone interactive program opens and the queue waits while it runs");
        var results = await task.WaitAsync(TimeSpan.FromSeconds(8));
        Assert(results.Count == 2 && results.All(r => r.State == InstallState.Succeeded) && File.ReadAllLines(marker).SequenceEqual(["opened", "closed"])
            && File.Exists(nextMarker), "closing the interactive program continues to the next item despite legacy D-drive settings and detection paths");
        Assert(updates.Any(p => p.Id == manual.Id && p.Phase == InstallPhase.WaitingForUser), "interactive execution reports the waiting-for-user phase");

        foreach (var mode in new[] { InstallDirectoryMode.Nsis, InstallDirectoryMode.InnoSetup, InstallDirectoryMode.MsiProperty })
        {
            var item = new BundleSoftware { Id = "directory", Arguments = mode == InstallDirectoryMode.Nsis ? "/S" : "",
                DefaultInstallDirectory = @"D:\", DirectoryMode = mode, DirectoryParameter = "INSTALLDIR", IsMsi = mode == InstallDirectoryMode.MsiProperty };
            var info = InstallationEngine.CreateStartInfo(item, Environment.ProcessPath!, root, false, discardDirectory: true);
            Assert(!info.Arguments.Contains("D:") && !info.Arguments.Contains("/DIR=") && !info.Arguments.Contains("INSTALLDIR="),
                $"{mode} uses the installer's own default directory without injecting an override");
            Assert(!InstallationEngine.CreateStartInfo(item, Environment.ProcessPath!, root, true).Arguments.Contains("D:"),
                $"{mode} also leaves the interactive installer's directory unchanged");
        }

        var installedMarker = Path.Combine(root, "installed.pid");
        File.WriteAllText(installedMarker + ".application", Path.Combine(fixtureDirectory, Path.GetFileName(Environment.ProcessPath!)));
        var launcher = new BundleSoftware { Id = Guid.NewGuid().ToString("N"), Name = "启动已安装软件", Installer = "setup.exe",
            Arguments = $"--process-fixture launch-installed 30000 \"{installedMarker}\"" };
        Process? application = null;
        try
        {
            results = await InstallationEngine.RunAsync([launcher, next], root, Path.Combine(root, "logs-installed-app"),
                resolveInstaller: _ => Environment.ProcessPath!, cancelCurrent: true, useInstallerDefaults: true).WaitAsync(TimeSpan.FromSeconds(8));
            clock.Restart();
            while (!File.Exists(installedMarker) && clock.Elapsed < TimeSpan.FromSeconds(8)) await Task.Delay(20);
            application = Process.GetProcessById(int.Parse(File.ReadAllText(installedMarker)));
            Assert(results.Count == 2 && results.All(r => r.State == InstallState.Succeeded) && !application.HasExited,
                "a newly installed application outside the staging directories neither blocks the queue nor gets killed on normal completion");
            Assert(File.ReadAllText(Path.Combine(root, "logs-installed-app", "session.log")).Contains("保留运行并继续队列"),
                "the log identifies the application left running after installer completion");
        }
        finally
        {
            if (application != null) { if (!application.HasExited) { application.Kill(); await application.WaitForExitAsync(); } application.Dispose(); }
        }
    }
}
