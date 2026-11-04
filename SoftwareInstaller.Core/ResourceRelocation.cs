using SoftwareInstaller.Models;

namespace SoftwareInstaller.Core;

public sealed record ResourcePathChange(string Scope, SoftwareItem Item, string? Installer, string? Resources, string Status, bool CanApply);

public static class ResourceRelocation
{
    public static int MissingCount(WorkspaceConfiguration configuration) => Entries(configuration).Count(p => Missing(p.Item));

    private static bool Missing(SoftwareItem item) => !File.Exists(item.FilePath)
        || (!string.IsNullOrWhiteSpace(item.ResourceDirectory) && !Directory.Exists(item.ResourceDirectory));

    private static IEnumerable<(string Scope, SoftwareItem Item)> Entries(WorkspaceConfiguration configuration) =>
        configuration.Software.Select(s => ("当前列表", s)).Concat(configuration.Schemes.SelectMany(p => p.Value.Software.Select(s => ("方案：" + p.Key, s))));

    public static string SuggestOldRoot(WorkspaceConfiguration configuration)
    {
        string? common = null;
        foreach (var (_, item) in Entries(configuration).Where(p => Missing(p.Item)))
        {
            var folder = !string.IsNullOrWhiteSpace(item.ResourceDirectory) ? item.ResourceDirectory : Path.GetDirectoryName(item.FilePath);
            if (string.IsNullOrWhiteSpace(folder) || !Path.IsPathFullyQualified(folder)) continue;
            folder = Path.GetFullPath(folder);
            if (common == null) { common = folder; continue; }
            while (!IsWithin(common, folder))
            {
                common = Path.GetDirectoryName(Path.TrimEndingDirectorySeparator(common));
                if (common == null) return "";
            }
        }
        return common ?? "";
    }

    private static bool IsWithin(string root, string path)
    {
        string relative = Path.GetRelativePath(root, path);
        return !Path.IsPathRooted(relative) && relative != ".." && !relative.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal);
    }

    public static List<ResourcePathChange> Prepare(WorkspaceConfiguration configuration, string oldRoot, string newRoot)
    {
        if (!Path.IsPathFullyQualified(oldRoot) || !Path.IsPathFullyQualified(newRoot)) throw new InvalidDataException("请填写原目录和新目录的完整路径。");
        oldRoot = Path.GetFullPath(oldRoot); newRoot = Path.GetFullPath(newRoot);
        if (!Directory.Exists(newRoot)) throw new DirectoryNotFoundException("新资源目录不存在，请重新选择。");
        var result = new List<ResourcePathChange>();
        foreach (var (scope, item) in Entries(configuration).Where(p => Missing(p.Item)))
        {
            var errors = new List<string>();
            string? Relocate(string? path, bool isDirectory)
            {
                if (isDirectory ? Directory.Exists(path) : File.Exists(path)) return path;
                if (string.IsNullOrWhiteSpace(path)) { if (!isDirectory) errors.Add("未设置安装文件"); return path; }
                if (!IsWithin(oldRoot, path)) { errors.Add(isDirectory ? "资源目录不在原目录内" : "安装文件不在原目录内"); return path; }
                string target = Path.GetFullPath(Path.GetRelativePath(oldRoot, path), newRoot);
                if (!(isDirectory ? Directory.Exists(target) : File.Exists(target))) errors.Add(isDirectory ? "新位置未找到资源目录" : "新位置未找到安装文件");
                return target;
            }
            string? installer = Relocate(item.FilePath, false), resources = Relocate(item.ResourceDirectory, true);
            if (errors.Count == 0 && !string.IsNullOrWhiteSpace(resources) && installer != null && !IsWithin(resources, installer))
                errors.Add("安装文件不在关联资源目录内，请单独编辑");
            bool changed = !string.Equals(item.FilePath, installer, StringComparison.OrdinalIgnoreCase) || !string.Equals(item.ResourceDirectory, resources, StringComparison.OrdinalIgnoreCase);
            bool valid = errors.Count == 0 && changed;
            result.Add(new(scope, item, installer, resources, valid ? "可修复" : string.Join("；", errors), valid));
        }
        return result;
    }

    public static int Apply(IEnumerable<ResourcePathChange> changes)
    {
        var selected = changes.Where(c => c.CanApply).ToList();
        // Recheck the entire selection before changing anything, in case sources moved after preview.
        if (selected.Any(c => !File.Exists(c.Installer) || (!string.IsNullOrWhiteSpace(c.Resources) && !Directory.Exists(c.Resources))))
            throw new IOException("预览中的安装源已发生变化，请重新预览后再应用。");
        foreach (var change in selected)
        {
            change.Item.Id = change.Item.Id;
            change.Item.FilePath = change.Installer; change.Item.ResourceDirectory = change.Resources;
        }
        return selected.Count;
    }
}
