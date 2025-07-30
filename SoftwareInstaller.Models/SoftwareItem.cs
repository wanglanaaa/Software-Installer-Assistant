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
        public string? FilePath { get; set; } // 用于存储安装程序的路径
        public string? SilentInstallArgs { get; set; } // 用于静默安装参数

        public SoftwareItem Clone()
        {
            return (SoftwareItem)this.MemberwiseClone();
        }
    }
}
