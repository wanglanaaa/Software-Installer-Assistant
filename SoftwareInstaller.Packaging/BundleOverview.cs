using SoftwareInstaller.Models;

namespace SoftwareInstaller.Packaging;

public sealed class BundleOverview
{
    public const long SpaceHeadroom = 100L * 1024 * 1024;
    public int SoftwareCount { get; private init; }
    public int SchemeCount { get; private init; }
    public int CategoryCount { get; private init; }
    public int FileCount { get; private init; }
    public long SourceBytes { get; private init; }
    public long RunnerBytes { get; private init; }
    public long OutputReferenceBytes => checked(SourceBytes + RunnerBytes + BundleReader.FooterSize);
    public long BuildFreeBytes => checked(SourceBytes * 2 + RunnerBytes + SpaceHeadroom);
    public long TargetTempBytes => checked(SourceBytes + SpaceHeadroom);
    public List<string> NetworkSoftware { get; private init; } = new();
    public List<string> InteractiveSoftware { get; private init; } = new();
    public List<string> DirectorySelectionSoftware { get; private init; } = new();

    public static BundleOverview Create(BundleBuildRequest request, IEnumerable<BundleSoftware>? descriptions = null)
    {
        var files = request.Software.SelectMany(BundleBuilder.GetSourceFiles).ToList();
        var items = (descriptions ?? request.Software.Select(BundleBuilder.DescribeForBundle)).ToList();
        return new()
        {
            SoftwareCount = request.Software.Count, SchemeCount = request.Schemes.Count,
            CategoryCount = CategoryNames.Normalize(request.Categories.Concat(request.Software.Select(s => s.Category))).Count,
            FileCount = files.Count, SourceBytes = files.Sum(f => new FileInfo(f.Source).Length), RunnerBytes = new FileInfo(request.RunnerPath).Length,
            NetworkSoftware = items.Where(s => s.RequiresNetwork).Select(s => s.Name).ToList(),
            InteractiveSoftware = items.Where(s => s.StartInteractively).Select(s => s.Name).ToList(),
            DirectorySelectionSoftware = items.Where(s => !string.IsNullOrWhiteSpace(s.DefaultInstallDirectory) && s.DirectoryMode == InstallDirectoryMode.Automatic).Select(s => s.Name).ToList()
        };
    }

    public static string FormatBytes(long bytes)
    {
        string[] units = ["B", "KiB", "MiB", "GiB", "TiB"];
        double value = bytes; int unit = 0;
        while (value >= 1024 && unit < units.Length - 1) { value /= 1024; unit++; }
        return $"{value:0.##} {units[unit]}";
    }
}
