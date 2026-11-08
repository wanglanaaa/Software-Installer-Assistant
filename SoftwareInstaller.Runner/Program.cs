using SoftwareInstaller.Packaging;

namespace SoftwareInstaller.Runner;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
#if NETFRAMEWORK
        // Only the minimum runtime prerequisite is checked; never scan the machine or installers.
        using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
        {
            if (!(key?.GetValue("Release") is int release) || release < 528040)
            {
                MessageBox.Show("此软件合集需要预先安装 .NET Framework 4.8。安装 .NET 9 不能替代此环境。", "缺少运行环境", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 1;
            }
        }
        EmbeddedDependencies.Register();
#endif
        return Run(args);
    }

    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    private static int Run(string[] args)
    {
        // Development measurement: load the actual UI/list, record readiness, never install or extract.
        if (args.Length == 2 && args[0] == "--measure-startup") return MeasureStartup(args[1]);
        // Diagnostic mode checks the package structure and extracts resources without starting installers.
        if ((args.Length == 2 || args.Length == 4 && args[2] == "--preview") && args[0] == "--verify-package")
        {
            var work = Path.Combine(Path.GetTempPath(), "SIA-verify-" + Guid.NewGuid().ToString("N"));
            try
            {
                using var package = BundleReader.OpenAsync(RunnerRuntime.ExecutablePath).GetAwaiter().GetResult();
                package.ExtractAsync(package.Manifest.Software, work).GetAwaiter().GetResult();
                if (args.Length == 4)
                {
                    ApplicationConfiguration.Initialize();
                    using var form = new InstallForm(RunnerRuntime.ExecutablePath) { ShowInTaskbar = false, Opacity = 0 };
                    form.Show();
                    var timer = System.Diagnostics.Stopwatch.StartNew();
                    while (!form.PackageReady && timer.Elapsed < TimeSpan.FromSeconds(30)) { Application.DoEvents(); Thread.Sleep(10); }
                    if (!form.PackageReady) throw new InvalidDataException("安装器预览未能加载安装清单。");
                    Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(args[3]))!);
                    using var bitmap = new Bitmap(form.Width, form.Height);
                    form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.Size)); bitmap.Save(args[3], System.Drawing.Imaging.ImageFormat.Png);
                    form.Close();
                }
                File.WriteAllText(args[1], RunnerRuntime.Json(new { success = true, package.Manifest.Name,
                    softwareCount = package.Manifest.Software.Count, schemeCount = package.Manifest.Schemes.Count,
                    runnerVersion = RunnerTemplate.CurrentVersion, logoWidth = AppBranding.Image.Width, logoHeight = AppBranding.Image.Height }));
                return 0;
            }
            catch (Exception ex) { File.WriteAllText(args[1], RunnerRuntime.Json(new { success = false, error = ex.Message })); return 1; }
            finally { if (Directory.Exists(work)) Directory.Delete(work, true); }
        }
        ApplicationConfiguration.Initialize();
        var selection = args.Length == 2 && args[0] == "--install" ? args[1].Split(',') : null;
        Application.Run(new InstallForm(RunnerRuntime.ExecutablePath, selection));
        return 0;
    }

    private static int MeasureStartup(string report)
    {
        var clock = System.Diagnostics.Stopwatch.StartNew();
        ApplicationConfiguration.Initialize();
        using var form = new InstallForm(RunnerRuntime.ExecutablePath) { ShowInTaskbar = false, Opacity = 0 };
        using var timer = new System.Windows.Forms.Timer { Interval = 10 };
        bool ready = false;
        timer.Tick += (_, _) =>
        {
            if (!form.PackageReady && clock.Elapsed < TimeSpan.FromSeconds(60)) return;
            ready = form.PackageReady;
            using var process = System.Diagnostics.Process.GetCurrentProcess();
            File.WriteAllText(report, RunnerRuntime.Json(new { ready, readyAfterMainMs = clock.Elapsed.TotalMilliseconds,
                workingSetBytes = process.WorkingSet64, privateBytes = process.PrivateMemorySize64 }));
            timer.Stop(); form.Close();
        };
        timer.Start(); Application.Run(form);
        return ready ? 0 : 1;
    }
}
