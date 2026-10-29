namespace SoftwareInstaller.Models;

public sealed class SchemeDefinition
{
    public List<SoftwareItem> Software { get; set; } = new();
    public List<string> Categories { get; set; } = new();
}

public static class CategoryNames
{
    public static List<string> Normalize(IEnumerable<string> names) => names.Where(n => !string.IsNullOrWhiteSpace(n))
        .Select(n => n.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
}
