namespace SoftwareInstaller.Models;

public static class UserDataPaths
{
    public static string? DataDirectoryOverride { get; set; }
    public static string DirectoryPath => DataDirectoryOverride ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SoftwareInstaller");

    public static string ReadPath(string fileName)
    {
        var current = Path.Combine(DirectoryPath, fileName);
        if (DataDirectoryOverride != null) return current;
        if (File.Exists(current)) return current;
        // Keep old portable files intact; the first save creates a per-user copy.
        var portable = Path.Combine(AppContext.BaseDirectory, fileName);
        if (File.Exists(portable)) return portable;
        return Path.GetFullPath(fileName);
    }

    public static void Write(string fileName, string json)
    {
        WriteFile(Path.Combine(DirectoryPath, fileName), json);
    }

    public static void WriteFile(string target, string json)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(target))!);
        var temporary = target + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try { File.WriteAllText(temporary, json); File.Move(temporary, target, true); }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }
}
