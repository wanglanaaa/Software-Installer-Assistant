namespace SoftwareInstaller.Packaging;

public sealed class InstallationPlan
{
    private readonly Dictionary<string, BundleSoftware> byId;
    public IReadOnlyList<BundleSoftware> Items { get; }

    private InstallationPlan(List<BundleSoftware> sorted)
    {
        Items = sorted.AsReadOnly();
        byId = sorted.ToDictionary(item => item.Id, StringComparer.Ordinal);
    }

    public static InstallationPlan Create(IEnumerable<BundleSoftware> software) => new(InstallationEngine.Sort(software));

    public InstallationPlan Select(IEnumerable<string> selectedIds)
    {
        var selected = selectedIds.ToHashSet(StringComparer.Ordinal);
        var pending = new Stack<string>(selected);
        while (pending.Count > 0)
        {
            var id = pending.Pop();
            if (!byId.TryGetValue(id, out var item)) throw new InvalidDataException("选择了未知的软件。");
            foreach (var dependency in item.DependencyIds)
                if (selected.Add(dependency)) pending.Push(dependency);
        }
        return new(Items.Where(item => selected.Contains(item.Id)).ToList());
    }
}
