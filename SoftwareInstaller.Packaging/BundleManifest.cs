using SoftwareInstaller.Models;

namespace SoftwareInstaller.Packaging;

public sealed class BundleManifest
{
    public int FormatVersion { get; set; } = 1;
    public string Name { get; set; } = "软件合集";
    public DateTimeOffset CreatedAt { get; set; } = DateTimeOffset.Now;
    public string DefaultScheme { get; set; } = "当前方案";
    public List<BundleSoftware> Software { get; set; } = new();
    public List<string> Categories { get; set; } = new();
    public Dictionary<string, List<string>> Schemes { get; set; } = new();
}

public sealed class BundleSoftware
{
    public string Id { get; set; } = "";
    public string Name { get; set; } = "";
    public string Version { get; set; } = "";
    public byte[]? IconPng { get; set; }
    public string Category { get; set; } = "";
    public string Installer { get; set; } = "";
    public string Arguments { get; set; } = "";
    public string? DefaultInstallDirectory { get; set; }
    public InstallDirectoryMode DirectoryMode { get; set; }
    public string? DirectoryParameter { get; set; }
    public bool IsMsi { get; set; }
    public bool StartInteractively { get; set; }
    public bool AllowInteractiveFallback { get; set; } = true;
    public bool RequiresNetwork { get; set; }
    public string InstallerType { get; set; } = "";
    public int Order { get; set; }
    public int TimeoutMinutes { get; set; } = 30;
    public List<int> SuccessCodes { get; set; } = new() { 0 };
    public List<int> RebootCodes { get; set; } = new() { 3010, 1641 };
    public string? DetectionFile { get; set; }
    public string? DetectionVersion { get; set; }
    public string? MsiProductCode { get; set; }
    public List<string> DependencyIds { get; set; } = new();
    public List<BundleFile> Files { get; set; } = new();
}

public sealed class BundleFile
{
    public string Path { get; set; } = "";
    public string Sha256 { get; set; } = "";
    public long Size { get; set; }
}

public sealed record BundleProgress(string Message, int Completed, int Total);
