using System.Text.Json;
using SoftwareInstaller.Models;

namespace SoftwareInstaller.Core;

public static class ConfigurationStore
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public static WorkspaceConfiguration Load(string path, out bool importedLegacy)
    {
        using var document = JsonDocument.Parse(File.ReadAllText(path));
        var root = document.RootElement;
        WorkspaceConfiguration configuration;
        importedLegacy = false;
        if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("Software", out _))
        {
            configuration = root.Deserialize<WorkspaceConfiguration>() ?? throw new InvalidDataException("配置文件为空。");
            if (configuration.FormatVersion != 1) throw new InvalidDataException("不支持的配置文件版本。");
        }
        else if (root.ValueKind == JsonValueKind.Array)
        {
            importedLegacy = true;
            configuration = root.GetArrayLength() > 0 && root[0].ValueKind == JsonValueKind.String
                ? new() { Categories = root.Deserialize<List<string>>()! }
                : new() { Software = root.Deserialize<List<SoftwareItem>>()! };
        }
        else if (root.ValueKind == JsonValueKind.Object)
        {
            importedLegacy = true;
            configuration = new() { Schemes = SchemeManager.LoadSchemes(path) };
        }
        else throw new InvalidDataException("请选择有效的软件安装配置 JSON 文件。");

        Validate(configuration);
        var baseDirectory = Path.GetDirectoryName(Path.GetFullPath(path))!;
        foreach (var item in configuration.Software.Concat(configuration.Schemes.Values.SelectMany(s => s.Software)))
        {
            if (!string.IsNullOrWhiteSpace(item.FilePath)) item.FilePath = Path.GetFullPath(item.FilePath, baseDirectory);
            if (!string.IsNullOrWhiteSpace(item.ResourceDirectory)) item.ResourceDirectory = Path.GetFullPath(item.ResourceDirectory, baseDirectory);
        }
        return configuration;
    }

    public static void Save(string path, WorkspaceConfiguration configuration)
    {
        Validate(configuration);
        UserDataPaths.WriteFile(path, Serialize(configuration));
    }

    public static string Serialize(WorkspaceConfiguration configuration) => JsonSerializer.Serialize(configuration, JsonOptions);

    private static void Validate(WorkspaceConfiguration configuration)
    {
        if (configuration.Software == null || configuration.Categories == null || configuration.Schemes == null)
            throw new InvalidDataException("配置缺少软件、分类或方案列表。");
        void CheckItems(List<SoftwareItem>? items)
        {
            if (items == null || items.Any(s => s == null || string.IsNullOrWhiteSpace(s.Name) || string.IsNullOrWhiteSpace(s.Id) || s.DependencyIds == null))
                throw new InvalidDataException("配置中存在无效的软件条目。");
            if (items.Select(s => s.Id).Distinct(StringComparer.Ordinal).Count() != items.Count)
                throw new InvalidDataException("配置中软件标识重复，请移除重复条目。");
        }
        CheckItems(configuration.Software);
        configuration.Categories = CategoryNames.Normalize(configuration.Categories.Concat(configuration.Software.Select(s => s.Category)));
        foreach (var (name, scheme) in configuration.Schemes)
        {
            if (string.IsNullOrWhiteSpace(name) || scheme == null || scheme.Categories == null)
                throw new InvalidDataException("配置中存在无效的安装方案。");
            CheckItems(scheme.Software);
            scheme.Categories = CategoryNames.Normalize(scheme.Categories.Concat(scheme.Software.Select(s => s.Category)));
        }
    }
}
