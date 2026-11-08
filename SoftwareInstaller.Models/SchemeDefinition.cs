namespace SoftwareInstaller.Models;

#if !NETFRAMEWORK
public sealed class SchemeDefinition
{
    public List<SoftwareItem> Software { get; set; } = new();
    public List<string> Categories { get; set; } = new();
}
#endif

public static class CategoryNames
{
    public static List<string> Normalize(IEnumerable<string> names) => names.Where(n => !string.IsNullOrWhiteSpace(n))
        .Select(n => n.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
}
