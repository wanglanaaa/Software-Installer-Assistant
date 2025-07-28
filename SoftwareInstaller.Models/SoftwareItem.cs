using System;

namespace SoftwareInstaller.Models
{
    public class SoftwareItem
    {
        public required string Name { get; set; }
        public required string Version { get; set; }
        public required string Size { get; set; }
        public required string Description { get; set; }
        public bool IsSelected { get; set; }
        public required string Category { get; set; }
        public string? FilePath { get; set; } // Added to store the path to the installer
        public string? SilentInstallArgs { get; set; } // Added for silent installation arguments

        public SoftwareItem Clone()
        {
            return (SoftwareItem)this.MemberwiseClone();
        }
    }
}
