using System.Text.Json;
using SoftwareInstaller.Models;

namespace SoftwareInstaller.Core;

public sealed class CategoryStore(string? storagePath = null)
{
    public List<string> Load(IEnumerable<string>? legacyCategories = null)
    {
        var path = storagePath ?? UserDataPaths.ReadPath("categories.json");
        return CategoryNames.Normalize(File.Exists(path)
            ? JsonSerializer.Deserialize<List<string>>(File.ReadAllText(path)) ?? new()
            : legacyCategories ?? []);
    }

    public void Save(IEnumerable<string> categories)
    {
        var json = JsonSerializer.Serialize(CategoryNames.Normalize(categories), new JsonSerializerOptions { WriteIndented = true });
        if (storagePath == null) UserDataPaths.Write("categories.json", json);
        else UserDataPaths.WriteFile(storagePath, json);
    }
}
