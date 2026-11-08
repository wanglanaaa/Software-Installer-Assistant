namespace SoftwareInstaller.Packaging;

public static class PackagePaths
{
    public static string Resolve(string root, string relative)
    {
        if (string.IsNullOrWhiteSpace(relative) || relative.Contains("\\") || relative.Contains(":") ||
            relative.Split('/').Any(s => s is "" or "." or ".." || s.EndsWith(" ", StringComparison.Ordinal) || s.EndsWith(".", StringComparison.Ordinal)) ||
            relative.IndexOfAny(['\0', '"', '<', '>', '|', '?', '*']) >= 0 || Path.IsPathRooted(relative))
            throw new InvalidDataException($"非法资源路径：{relative}");
        var path = Path.GetFullPath(Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar)));
        var prefix = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("资源路径超出工作目录。");
        return path;
    }

    public static bool IsReparsePoint(string path) => (File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0;
}
