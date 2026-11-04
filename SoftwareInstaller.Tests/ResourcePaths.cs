using System.Reflection;
using System.Text.Json;
using SoftwareInstaller.Core;
using SoftwareInstaller.Models;
using SoftwareInstaller.UI;

namespace SoftwareInstaller.Tests;

internal static partial class Program
{
    private static void TestResourcePaths()
    {
        SoftwareItem Source(string name, string path, string? resources = null) => new() { Name = name, FilePath = path, ResourceDirectory = resources,
            Category = "工具", Description = "说明", Version = "1", Size = "", IsSelected = true, DefaultInstallDirectory = "D:\\Apps", InstallOrder = 7 };
        var original = Path.Combine(root, "original"); var moved = Path.Combine(root, "moved");
        var resources = Path.Combine(original, "安装源", "中文 软件"); Directory.CreateDirectory(resources);
        var source = Path.Combine(resources, "setup.msi"); File.WriteAllText(source, "safe test data");
        var item = Source("办公助手", source, resources);
        string id = item.Id;
        var configuration = new WorkspaceConfiguration { Software = [item], Categories = ["工具", "空分类"],
            Schemes = new() { ["常用"] = new() { Software = [item.Clone()], Categories = ["工具", "空分类"] } } };
        string originalText = ConfigurationStore.Serialize(configuration);
        var json = Path.Combine(original, "配置", "安装配置.json"); ConfigurationStore.Save(json, configuration);
        using (var document = JsonDocument.Parse(File.ReadAllText(json)))
        {
            var stored = document.RootElement.GetProperty("Software")[0];
            Assert(!Path.IsPathRooted(stored.GetProperty("FilePath").GetString()!) && !Path.IsPathRooted(stored.GetProperty("ResourceDirectory").GetString()!),
                "saving uses relative installer and resource paths, including sibling folders");
            Assert(stored.GetProperty("DefaultInstallDirectory").GetString() == "D:\\Apps", "saving relative source paths does not alter the target installation directory");
        }
        Assert(ConfigurationStore.Serialize(configuration) == originalText, "saving does not replace live absolute paths or change legacy identifiers");
        // Copy the complete tree to simulate moving a library without modifying any user files.
        foreach (var file in Directory.EnumerateFiles(original, "*", SearchOption.AllDirectories))
        {
            var destination = Path.Combine(moved, Path.GetRelativePath(original, file)); Directory.CreateDirectory(Path.GetDirectoryName(destination)!); File.Copy(file, destination);
        }
        var loaded = ConfigurationStore.Load(Path.Combine(moved, "配置", "安装配置.json"), out _);
        string movedFile = Path.Combine(moved, "安装源", "中文 软件", "setup.msi");
        Assert(loaded.Software[0].FilePath == movedFile && loaded.Schemes["常用"].Software[0].FilePath == movedFile && loaded.Software[0].Id == id && ResourceRelocation.MissingCount(loaded) == 0,
            "moving the entire library restores current and saved-scheme sources relative to the JSON file");
        var anotherJson = Path.Combine(root, "另存位置", "renamed.json"); ConfigurationStore.Save(anotherJson, loaded);
        Assert(ConfigurationStore.Load(anotherJson, out _).Software[0].FilePath == movedFile, "Save As recalculates relative paths for the new configuration directory");
        var crossDrive = Source("其他盘", "Z:\\媒体库\\setup.exe");
        var crossJson = Path.Combine(root, "cross-drive.json"); ConfigurationStore.Save(crossJson, new() { Software = [crossDrive] });
        using (var document = JsonDocument.Parse(File.ReadAllText(crossJson)))
            Assert(Path.IsPathFullyQualified(document.RootElement.GetProperty("Software")[0].GetProperty("FilePath").GetString()!), "sources on another drive remain absolute instead of producing unusable relative paths");

        var oldRoot = Path.Combine(root, "missing-old"); var newRoot = Path.Combine(root, "new-library");
        Directory.CreateDirectory(Path.Combine(newRoot, "App"));
        File.WriteAllText(Path.Combine(newRoot, "App", "setup.exe"), "safe"); File.WriteAllText(Path.Combine(newRoot, "other.exe"), "safe");
        var missing = Source("待修复", Path.Combine(oldRoot, "App", "setup.exe"), Path.Combine(oldRoot, "App"));
        var absent = Source("未找到", Path.Combine(oldRoot, "unmatched", "setup.exe"));
        var outside = Source("其他库", Path.Combine(oldRoot + "-other", "other.exe"));
        var valid = Source("已有路径", movedFile);
        var repair = new WorkspaceConfiguration { Software = [missing, absent, outside, valid], Categories = ["工具", "空分类"],
            Schemes = new() { ["常用"] = new() { Software = [missing.Clone()] } } };
        string beforePreview = ConfigurationStore.Serialize(repair);
        var changes = ResourceRelocation.Prepare(repair, oldRoot, newRoot);
        Assert(changes.Count == 4 && changes.Count(c => c.CanApply) == 2 && ConfigurationStore.Serialize(repair) == beforePreview,
            "repair preview includes saved schemes, excludes valid sources and changes nothing before applying");
        Assert(changes.Where(c => c.Item.Name is "未找到" or "其他库").All(c => !c.CanApply), "same-named files and partial directory-prefix matches are not guessed as replacements");
        ResourceRelocation.Apply(changes);
        Assert(missing.FilePath == Path.Combine(newRoot, "App", "setup.exe") && missing.ResourceDirectory == Path.Combine(newRoot, "App") && repair.Schemes["常用"].Software[0].FilePath == missing.FilePath
            && missing.IsSelected && missing.InstallOrder == 7 && repair.Categories.SequenceEqual(["工具", "空分类"]) && valid.FilePath == movedFile,
            "applying repairs both paths and scheme snapshots while preserving selections, order, categories and valid sources");
        var stale = Source("预览后变化", Path.Combine(oldRoot, "other.exe"));
        var stalePlan = ResourceRelocation.Prepare(new() { Software = [stale] }, oldRoot, newRoot);
        File.Move(Path.Combine(newRoot, "other.exe"), Path.Combine(newRoot, "moved-away.exe"));
        bool rejected = false;
        try { ResourceRelocation.Apply(stalePlan); } catch (IOException) { rejected = true; }
        Assert(rejected && stale.FilePath == Path.Combine(oldRoot, "other.exe"), "sources disappearing after preview reject the repair without partial changes");
        RenderResourceRepair(repair, oldRoot, newRoot);
    }

    private static void RenderResourceRepair(WorkspaceConfiguration configuration, string oldRoot, string newRoot)
    {
        Application.EnableVisualStyles();
        T Field<T>(object form, string name) => (T)form.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
        var previewConfiguration = new WorkspaceConfiguration { Software = configuration.Software.Select(s => s.Clone()).ToList(), Categories = new(configuration.Categories) };
        previewConfiguration.Software[0].FilePath = Path.Combine(oldRoot, "App", "setup.exe");
        previewConfiguration.Software[0].ResourceDirectory = Path.Combine(oldRoot, "App");
        string snapshot = ConfigurationStore.Serialize(previewConfiguration);
        using var repair = new ResourceRepairForm(previewConfiguration) { Opacity = 0, ShowInTaskbar = false };
        repair.Show(); Application.DoEvents();
        Field<TextBox>(repair, "oldRoot").Text = oldRoot; Field<TextBox>(repair, "newRoot").Text = newRoot;
        Field<Button>(repair, "preview").PerformClick(); Application.DoEvents();
        var grid = Field<DataGridView>(repair, "grid");
        Assert(grid.Rows.Count == 3 && grid.Rows.Cast<DataGridViewRow>().Count(r => r.Cells[0].ReadOnly) == 2 && Field<Button>(repair, "apply").Enabled,
            "repair UI preselects valid replacements and prevents selection of unresolved records");
        using (var picture = new Bitmap(repair.Width, repair.Height)) { repair.DrawToBitmap(picture, new Rectangle(Point.Empty, repair.Size)); picture.Save(Path.Combine(root, "preview-resource-repair.png")); }
        repair.Size = repair.MinimumSize; Application.DoEvents();
        using (var picture = new Bitmap(repair.Width, repair.Height)) { repair.DrawToBitmap(picture, new Rectangle(Point.Empty, repair.Size)); picture.Save(Path.Combine(root, "preview-resource-repair-minimum.png")); }
        Field<TextBox>(repair, "newRoot").Text = "";
        Assert(!Field<Button>(repair, "apply").Enabled && grid.Rows.Count == 0, "changing a directory invalidates an old repair preview");
        Field<Button>(repair, "preview").PerformClick();
        Assert(Field<Label>(repair, "status").Text.Contains("完整路径") && ConfigurationStore.Serialize(previewConfiguration) == snapshot,
            "invalid roots are explained in the UI and cancelling or previewing never changes the configuration");
        Field<TextBox>(repair, "newRoot").Text = newRoot; Field<Button>(repair, "preview").PerformClick(); Field<Button>(repair, "apply").PerformClick();
        Assert(repair.DialogResult == DialogResult.OK && repair.AppliedCount == 1 && previewConfiguration.Software[0].FilePath == Path.Combine(newRoot, "App", "setup.exe"),
            "repair UI applies only confirmed valid changes");

        using var main = new MainForm { Opacity = 0, ShowInTaskbar = false };
        main.Show(); main.Size = main.MinimumSize; Application.DoEvents();
        using var image = new Bitmap(main.Width, main.Height); main.DrawToBitmap(image, new Rectangle(Point.Empty, main.Size)); image.Save(Path.Combine(root, "preview-resource-entry.png"));
    }
}
