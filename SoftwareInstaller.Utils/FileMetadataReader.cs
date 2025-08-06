using System.Diagnostics;

namespace SoftwareInstaller.Utils
{
    public static class FileMetadataReader
    {
        public static string? GetProductName(string filePath)
        {
            try
            {
                var versionInfo = FileVersionInfo.GetVersionInfo(filePath);
                return versionInfo.ProductName;
            }
            catch
            {
                return null;
            }
        }
    }
}
