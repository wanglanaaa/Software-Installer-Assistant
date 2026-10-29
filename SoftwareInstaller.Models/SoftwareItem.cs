using System;

namespace SoftwareInstaller.Models
{
    public class SoftwareItem
    {
        private string? id;
        public string Id
        {
            // Legacy lists and scheme snapshots must resolve to the same identifier.
            get => id ??= Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(
                System.Text.Encoding.UTF8.GetBytes($"{Name}|{Version}|{FilePath}".ToUpperInvariant())))[..32].ToLowerInvariant();
            set => id = value;
        }
        public required string Name { get; set; }
        public required string Version { get; set; }
        public required string Size { get; set; }
        public required string Description { get; set; }
        public bool IsSelected { get; set; }
        public required string Category { get; set; }
        public string? FilePath { get; set; } // 用于存储安装程序的路径
        public string? SilentInstallArgs { get; set; } // 用于静默安装参数
        public string? DefaultInstallDirectory { get; set; }
        public InstallDirectoryMode DirectoryMode { get; set; } = InstallDirectoryMode.Automatic;
        public string? DirectoryParameter { get; set; }
        public string? VerifiedSha256 { get; set; }
        public string? ResourceDirectory { get; set; }
        public bool RequiresNetwork { get; set; }
        public bool AllowInteractiveFallback { get; set; } = true;
        public int InstallOrder { get; set; }
        public int TimeoutMinutes { get; set; } = 30;
        public string SuccessCodes { get; set; } = "0";
        public string RebootCodes { get; set; } = "3010,1641";
        public string? DetectionFile { get; set; }
        public string? DetectionVersion { get; set; }
        public string? MsiProductCode { get; set; }
        public List<string> DependencyIds { get; set; } = new();

        public SoftwareItem Clone()
        {
            var copy = (SoftwareItem)MemberwiseClone();
            copy.DependencyIds = new List<string>(DependencyIds);
            return copy;
        }
    }
}
