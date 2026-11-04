using System.Diagnostics;
using SoftwareInstaller.Models;

namespace SoftwareInstaller.Core;

public sealed record ImportPreparation(List<SoftwareItem> Items, int Duplicates, List<string> Errors);

public static class SoftwareImport
{
    public static ImportPreparation Prepare(IEnumerable<string> paths, IEnumerable<SoftwareItem> existing, string category)
    {
        var seen = existing.Where(s => !string.IsNullOrWhiteSpace(s.FilePath))
            .Select(s => Path.GetFullPath(s.FilePath!)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var items = new List<SoftwareItem>(); var errors = new List<string>(); int duplicates = 0;
        foreach (var input in paths)
        {
            try
            {
                var path = Path.GetFullPath(input);
                if (!new[] { ".exe", ".msi" }.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase) || !File.Exists(path))
                    throw new InvalidDataException("请选择存在的 EXE 或 MSI 安装文件。");
                if (!seen.Add(path)) { duplicates++; continue; }
                var item = new SoftwareItem { Id = Guid.NewGuid().ToString("N"), Name = Path.GetFileNameWithoutExtension(path),
                    FilePath = path, Version = "", Description = "", Size = $"{new FileInfo(path).Length / 1048576d:F2} MB",
                    Category = string.IsNullOrWhiteSpace(category) ? "未分类" : category, IsSelected = true };
                try
                {
                    var info = FileVersionInfo.GetVersionInfo(path);
                    item.Version = info.FileVersion?.Trim() ?? ""; item.Description = info.FileDescription?.Trim() ?? "";
                }
                catch (Exception ex) when (ex is IOException or System.ComponentModel.Win32Exception) { /* Filename remains usable without metadata. */ }
                items.Add(item);
            }
            catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or ArgumentException)
            { errors.Add($"{Path.GetFileName(input)}：{ex.Message}"); }
        }
        return new(items, duplicates, errors);
    }
}
