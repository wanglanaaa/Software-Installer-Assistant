using SoftwareInstaller.Models;

namespace SoftwareInstaller.Core;

public sealed class SchemeHandler
{
    private readonly Dictionary<string, SchemeDefinition> schemes;
    private readonly string? storagePath;
    private readonly bool useDefaultStorage;
    public SchemeHandler(string? storagePath = null, bool useDefaultStorage = true)
    {
        this.storagePath = storagePath;
        this.useDefaultStorage = useDefaultStorage;
        schemes = useDefaultStorage || storagePath != null ? SchemeManager.LoadSchemes(storagePath) : new();
    }
    public IEnumerable<string> GetSchemeNames() => schemes.Keys.OrderBy(k => k);
    public bool SchemeExists(string name) => schemes.ContainsKey(name);
    public List<SoftwareItem>? GetScheme(string name) => schemes.TryGetValue(name, out var scheme) ? scheme.Software.Select(s => s.Clone()).ToList() : null;
    public List<string> GetCategories(string name) => schemes.TryGetValue(name, out var scheme) ? new(scheme.Categories) : new();
    public void SaveScheme(string name, List<SoftwareItem> items, IEnumerable<string>? categories = null)
    {
        var updated = new Dictionary<string, SchemeDefinition>(schemes);
        updated[name] = new() { Software = items.Select(s => s.Clone()).ToList(), Categories = CategoryNames.Normalize((categories ?? []).Concat(items.Select(s => s.Category))) };
        if (useDefaultStorage || storagePath != null) SchemeManager.SaveSchemes(updated, storagePath);
        schemes[name] = updated[name];
    }
    public void DeleteScheme(string name)
    {
        if (!schemes.ContainsKey(name)) return;
        var updated = new Dictionary<string, SchemeDefinition>(schemes);
        updated.Remove(name);
        if (useDefaultStorage || storagePath != null) SchemeManager.SaveSchemes(updated, storagePath);
        schemes.Remove(name);
    }

    public Dictionary<string, SchemeDefinition> GetDefinitions() => schemes.ToDictionary(p => p.Key,
        p => new SchemeDefinition { Software = p.Value.Software.Select(s => s.Clone()).ToList(), Categories = new(p.Value.Categories) });

    public void ReplaceSchemes(Dictionary<string, SchemeDefinition> definitions)
    {
        var copies = definitions.ToDictionary(p => p.Key,
            p => new SchemeDefinition { Software = p.Value.Software.Select(s => s.Clone()).ToList(), Categories = new(p.Value.Categories) });
        schemes.Clear();
        foreach (var pair in copies) schemes.Add(pair.Key, pair.Value);
    }
}
