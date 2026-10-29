using System.Text.Json;
using SoftwareInstaller.Packaging;

namespace SoftwareInstaller.Runner;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        // Diagnostic mode verifies and extracts resources without starting any installers.
        if (args.Length == 2 && args[0] == "--verify-package")
        {
            var work = Path.Combine(Path.GetTempPath(), "SIA-verify-" + Guid.NewGuid().ToString("N"));
            try
            {
                using var package = BundleReader.OpenAsync(Environment.ProcessPath!).GetAwaiter().GetResult();
                package.ExtractAsync(package.Manifest.Software, work).GetAwaiter().GetResult();
                File.WriteAllText(args[1], JsonSerializer.Serialize(new { success = true, package.Manifest.Name,
                    softwareCount = package.Manifest.Software.Count, schemeCount = package.Manifest.Schemes.Count }));
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
