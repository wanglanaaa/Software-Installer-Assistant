using System.Runtime.InteropServices;
using System.Text;

namespace SoftwareInstaller.Models;

public enum InstallerKind { Unknown, Msi, InnoSetup, Nsis }

public sealed record InstallerProfile(InstallerKind Kind, string Name, string SilentArguments,
    InstallDirectoryMode DirectoryMode = InstallDirectoryMode.Automatic, string? DirectoryParameter = null);

public static class InstallerInspector
{
    public static readonly InstallerProfile Unknown = new(InstallerKind.Unknown, "未识别（使用安装界面）", "");

    public static InstallerProfile Inspect(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return Unknown;
        if (string.Equals(Path.GetExtension(path), ".msi", StringComparison.OrdinalIgnoreCase))
        {
            string? property = OperatingSystem.IsWindows() && File.Exists(path) ? ReadMsiDirectoryProperty(path) : null;
            return new(InstallerKind.Msi, "MSI（自动使用标准参数）", "", property == null ? InstallDirectoryMode.Automatic : InstallDirectoryMode.MsiProperty, property);
        }
        if (!File.Exists(path)) return Unknown;
        try
        {
            using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            int size = (int)Math.Min(file.Length, 2 * 1024 * 1024);
            var bytes = new byte[size]; file.ReadExactly(bytes);
            var profile = InspectBytes(bytes);
            if (profile.Kind != InstallerKind.Unknown || file.Length <= size) return profile;
            file.Position = Math.Max(size, file.Length - size);
            bytes = new byte[file.Length - file.Position]; file.ReadExactly(bytes);
            return InspectBytes(bytes);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { return Unknown; }
    }

    private static InstallerProfile InspectBytes(byte[] bytes)
    {
        var ascii = Encoding.Latin1.GetString(bytes);
        if (ascii.Contains("Inno Setup Setup Data", StringComparison.Ordinal) ||
            ascii.Contains("Inno Setup Setup Loader", StringComparison.Ordinal))
            return new(InstallerKind.InnoSetup, "Inno Setup", "/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-", InstallDirectoryMode.InnoSetup);
        if (ascii.Contains("NullsoftInst", StringComparison.Ordinal))
            return new(InstallerKind.Nsis, "NSIS", "/S", InstallDirectoryMode.Nsis);
        return Unknown;
    }

    private static string? ReadMsiDirectoryProperty(string path)
    {
        uint database = 0;
        try
        {
            if (MsiOpenDatabase(path, IntPtr.Zero, out database) != 0) return null;
            var directories = Query(database, "SELECT `Directory` FROM `Directory`");
            var wix = Query(database, "SELECT `Value` FROM `Property` WHERE `Property`='WIXUI_INSTALLDIR'").FirstOrDefault();
            if (wix != null && directories.Contains(wix) && IsPublicProperty(wix)) return wix;
            var common = new[] { "INSTALLFOLDER", "INSTALLDIR", "APPLICATIONFOLDER", "APPDIR" }.Where(directories.Contains).ToList();
            return common.Count == 1 ? common[0] : null;
        }
        catch (Exception ex) when (ex is DllNotFoundException or EntryPointNotFoundException) { return null; }
        finally { if (database != 0) MsiCloseHandle(database); }
    }

    private static bool IsPublicProperty(string value) => System.Text.RegularExpressions.Regex.IsMatch(value, @"^[A-Z_][A-Z0-9_]*$");

    private static HashSet<string> Query(uint database, string sql)
    {
        var values = new HashSet<string>(StringComparer.Ordinal);
        if (MsiDatabaseOpenView(database, sql, out uint view) != 0) return values;
        try
        {
            if (MsiViewExecute(view, 0) != 0) return values;
            while (MsiViewFetch(view, out uint record) == 0)
            {
                try
                {
                    var value = new StringBuilder(1024); uint length = 1024;
                    if (MsiRecordGetString(record, 1, value, ref length) == 0) values.Add(value.ToString());
                }
                finally { MsiCloseHandle(record); }
            }
        }
        finally { MsiCloseHandle(view); }
        return values;
    }

    [DllImport("msi.dll", EntryPoint = "MsiOpenDatabaseW", CharSet = CharSet.Unicode)]
    private static extern uint MsiOpenDatabase(string path, IntPtr mode, out uint database);
    [DllImport("msi.dll", EntryPoint = "MsiDatabaseOpenViewW", CharSet = CharSet.Unicode)]
    private static extern uint MsiDatabaseOpenView(uint database, string query, out uint view);
    [DllImport("msi.dll")] private static extern uint MsiViewExecute(uint view, uint record);
    [DllImport("msi.dll")] private static extern uint MsiViewFetch(uint view, out uint record);
    [DllImport("msi.dll", EntryPoint = "MsiRecordGetStringW", CharSet = CharSet.Unicode)]
    private static extern uint MsiRecordGetString(uint record, uint field, StringBuilder value, ref uint length);
    [DllImport("msi.dll")] private static extern uint MsiCloseHandle(uint handle);
}
