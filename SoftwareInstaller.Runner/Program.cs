using System.Text.Json;
using SoftwareInstaller.Packaging;

namespace SoftwareInstaller.Runner;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        // Diagnostic mode checks the package structure and extracts resources without starting installers.
        if ((args.Length == 2 || args.Length == 4 && args[2] == "--preview") && args[0] == "--verify-package")
        {
            var work = Path.Combine(Path.GetTempPath(), "SIA-verify-" + Guid.NewGuid().ToString("N"));
            try
            {
                using var package = BundleReader.OpenAsync(Environment.ProcessPath!).GetAwaiter().GetResult();
                package.ExtractAsync(package.Manifest.Software, work).GetAwaiter().GetResult();
                if (args.Length == 4)
                {
                    ApplicationConfiguration.Initialize();
                    using var form = new InstallForm(Environment.ProcessPath!) { ShowInTaskbar = false, Opacity = 0 };
                    form.Show();
                    var timer = System.Diagnostics.Stopwatch.StartNew();
                    while (!form.PackageReady && timer.Elapsed < TimeSpan.FromSeconds(30)) { Application.DoEvents(); Thread.Sleep(10); }
                    if (!form.PackageReady) throw new InvalidDataException("安装器预览未能加载安装清单。");
                    Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(args[3]))!);
                    using var bitmap = new Bitmap(form.Width, form.Height);
                    form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.Size)); bitmap.Save(args[3], System.Drawing.Imaging.ImageFormat.Png);
                    form.Close();
                }
                File.WriteAllText(args[1], JsonSerializer.Serialize(new { success = true, package.Manifest.Name,
                    softwareCount = package.Manifest.Software.Count, schemeCount = package.Manifest.Schemes.Count,
                    runnerVersion = RunnerTemplate.CurrentVersion, logoWidth = AppBranding.Image.Width, logoHeight = AppBranding.Image.Height }));
                return 0;
            }
            catch (Exception ex) { File.WriteAllText(args[1], JsonSerializer.Serialize(new { success = false, error = ex.Message })); return 1; }
            finally { if (Directory.Exists(work)) Directory.Delete(work, true); }
        }
        ApplicationConfiguration.Initialize();
        var selection = args.Length == 2 && args[0] == "--install" ? args[1].Split(',') : null;
        Application.Run(new InstallForm(Environment.ProcessPath!, selection));
        return 0;
    }
}
