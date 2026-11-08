using System.Text.RegularExpressions;

namespace SoftwareInstaller.Models;

public enum InstallDirectoryMode { InstallerDefault, InnoSetup, Nsis, MsiProperty, Custom, Automatic }

public static class InstallDirectoryOptions
{
    public const string Placeholder = "{InstallDir}";

    public static string Normalize(string? value)
    {
        var path = (value ?? "").Trim().Replace('/', '\\');
        if (Regex.IsMatch(path, @"^[A-Za-z]:$")) path += "\\";
        if (path.Length > 3) path = path.TrimEnd('\\');
        return path;
    }

    public static void Validate(string? value, InstallDirectoryMode mode, string? parameter, bool isMsi)
    {
        if (!Enum.IsDefined(typeof(InstallDirectoryMode), mode)) throw new InvalidDataException("安装目录参数类型无效。");
        var path = Normalize(value);
        if (path.Length == 0)
        {
            return;
        }
        if (mode == InstallDirectoryMode.InstallerDefault) throw new InvalidDataException("设置目录后，请选择安装器支持的目录参数类型。");
        if (!Regex.IsMatch(path, @"^[A-Za-z]:\\") || path.Substring(3).IndexOfAny(['\0', '\r', '\n', '"', '<', '>', '|', '?', '*', ':']) >= 0 ||
            path.Split('\\').Skip(1).Any(p => p is "." or ".." || p.EndsWith(" ", StringComparison.Ordinal) || p.EndsWith(".", StringComparison.Ordinal)))
            throw new InvalidDataException("安装目录须为本地磁盘绝对路径，例如 D:\\Apps\\软件名；输入 D: 按 D:\\ 处理。");
        if (isMsi && mode is InstallDirectoryMode.InnoSetup or InstallDirectoryMode.Nsis)
            throw new InvalidDataException("MSI 安装器请选择 MSI 属性或自定义参数。");
        if (!isMsi && mode == InstallDirectoryMode.MsiProperty)
            throw new InvalidDataException("MSI 属性模式只适用于 MSI 安装文件。");
        if (mode == InstallDirectoryMode.MsiProperty && !Regex.IsMatch(parameter ?? "", @"^[A-Z_][A-Z0-9_]*$"))
            throw new InvalidDataException("请填写厂商支持的 MSI 目录属性名，例如 INSTALLDIR；不同软件可能使用不同属性。");
        if (mode == InstallDirectoryMode.Custom && (string.IsNullOrWhiteSpace(parameter) || !parameter!.Contains(Placeholder) || parameter.IndexOfAny(['\0', '\r', '\n']) >= 0))
            throw new InvalidDataException("自定义目录参数必须包含 {InstallDir}，例如 --install-dir=\"{InstallDir}\"。");
    }

    public static string BuildArguments(string? arguments, string? directory, InstallDirectoryMode mode, string? parameter, bool isMsi)
    {
        Validate(directory, mode, parameter, isMsi);
        var path = Normalize(directory);
        var args = arguments?.Trim() ?? "";
        if (path.Length == 0)
        {
            if (args.Contains(Placeholder)) throw new InvalidDataException("参数中使用了 {InstallDir}，但未填写默认安装目录。");
            return args;
        }
        if (mode == InstallDirectoryMode.Automatic) return args;
        if (args.Contains(Placeholder)) throw new InvalidDataException("请将目录参数移到目录参数模板中，避免与静默参数重复。");
        string? duplicatePattern = mode switch
        {
            InstallDirectoryMode.Nsis => @"(?:^|\s)/D=",
            InstallDirectoryMode.InnoSetup => @"(?:^|\s)/DIR=",
            InstallDirectoryMode.MsiProperty => @"(?:^|\s)" + Regex.Escape(parameter!) + "=",
            _ => null
        };
        if (duplicatePattern != null && Regex.IsMatch(args, duplicatePattern, RegexOptions.IgnoreCase))
            throw new InvalidDataException("静默参数中已存在目录设置，请移除旧目录参数，统一使用默认安装目录配置。");
        var extra = mode switch
        {
            InstallDirectoryMode.InnoSetup => "/DIR=" + (path.Any(char.IsWhiteSpace) ? Quote(path) : path),
            // NSIS consumes the rest of the command line; /D must be last and unquoted.
            InstallDirectoryMode.Nsis => "/D=" + path,
            InstallDirectoryMode.MsiProperty => parameter + "=" + Quote(path),
            InstallDirectoryMode.Custom => parameter!.Replace("\"" + Placeholder + "\"", Quote(path)).Replace(Placeholder, path),
            _ => ""
        };
        return string.IsNullOrWhiteSpace(args) ? extra : args + " " + extra;
    }

#if !NETFRAMEWORK
    public static string BuildArguments(SoftwareItem item, string? arguments = null)
    {
        var profile = item.DirectoryMode == InstallDirectoryMode.Automatic && !string.IsNullOrWhiteSpace(item.DefaultInstallDirectory) ? InstallerInspector.Inspect(item.FilePath) : null;
        return BuildArguments(arguments ?? item.SilentInstallArgs, item.DefaultInstallDirectory,
            profile?.DirectoryMode ?? item.DirectoryMode, profile?.DirectoryParameter ?? item.DirectoryParameter,
            string.Equals(Path.GetExtension(item.FilePath), ".msi", StringComparison.OrdinalIgnoreCase));
    }
#endif

    public static void CheckTargetDrive(string? directory)
    {
        var path = Normalize(directory);
        if (path.Length > 0 && !Directory.Exists(Path.GetPathRoot(path)))
            throw new IOException($"目标磁盘 {Path.GetPathRoot(path)} 不存在或不可访问，未启动安装程序。请修改默认安装目录。");
    }

    public static string ExpandDetectionPath(string path, string? directory)
    {
        if (path.Contains(Placeholder) && string.IsNullOrWhiteSpace(directory)) throw new InvalidDataException("检测路径使用了 {InstallDir}，但未填写默认安装目录。");
        return Environment.ExpandEnvironmentVariables(path.Replace(Placeholder, Normalize(directory).TrimEnd('\\')));
    }

    public static string Quote(string value)
    {
        var result = new System.Text.StringBuilder("\"");
        int slashes = 0;
        foreach (char c in value)
        {
            if (c == '\\') { slashes++; continue; }
            result.Append('\\', c == '"' ? slashes * 2 + 1 : slashes);
            result.Append(c); slashes = 0;
        }
        return result.Append('\\', slashes * 2).Append('"').ToString();
    }
}
