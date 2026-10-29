using System.Text.Json;
using SoftwareInstaller.Models;

namespace SoftwareInstaller.Core;

public static class SchemeManager
{
    private sealed class SchemeFile
    {
        public int FormatVersion { get; set; } = 2;
        public Dictionary<string, SchemeDefinition> Schemes { get; set; } = new();
    }

    public static Dictionary<string, SchemeDefinition> LoadSchemes(string? storagePath = null)
    {
        var path = storagePath ?? UserDataPaths.ReadPath("schemes.json");
        if (!File.Exists(path)) return new();
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        Dictionary<string, SchemeDefinition> schemes;
        if (document.RootElement.TryGetProperty("FormatVersion", out var version) && version.ValueKind == JsonValueKind.Number)
        {
            if (version.GetInt32() != 2) throw new InvalidDataException("不支持的安装方案版本。");
            schemes = document.RootElement.Deserialize<SchemeFile>()?.Schemes ?? new();
        }
        else
        {
            // v1 stored only software lists; recover non-empty categories without discarding old schemes.
            var legacy = document.RootElement.Deserialize<Dictionary<string, List<SoftwareItem>>>() ?? new();
            schemes = legacy.ToDictionary(p => p.Key, p => new SchemeDefinition { Software = p.Value, Categories = CategoryNames.Normalize(p.Value.Select(s => s.Category)) });
        }
        foreach (var scheme in schemes.Values)
            scheme.Categories = CategoryNames.Normalize(scheme.Categories.Concat(scheme.Software.Select(s => s.Category)));
        return schemes;
    }

    public static void SaveSchemes(Dictionary<string, SchemeDefinition> schemes, string? storagePath = null)
    {
        var json = JsonSerializer.Serialize(new SchemeFile { Schemes = schemes }, new JsonSerializerOptions { WriteIndented = true });
        if (storagePath == null) UserDataPaths.Write("schemes.json", json);
        else UserDataPaths.WriteFile(storagePath, json);
    }
}
