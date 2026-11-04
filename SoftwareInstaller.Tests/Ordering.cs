using System.Reflection;
using SoftwareInstaller.Core;
using SoftwareInstaller.Models;
using SoftwareInstaller.Packaging;
using SoftwareInstaller.UI;

namespace SoftwareInstaller.Tests;

internal static partial class Program
{
    private static void TestManualOrdering()
    {
        Application.EnableVisualStyles();
        using var main = new MainForm { Opacity = 0, ShowInTaskbar = false };
        T Field<T>(string name) => (T)typeof(MainForm).GetField(name, BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(main)!;
        object? Call(string name, params object[] arguments) => typeof(MainForm).GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance)!.Invoke(main, arguments);
        var manager = Field<SoftwareManager>("_softwareManager");
        SoftwareItem Item(string id, string category, bool check) => new() { Id = id, Name = id, Category = category, IsSelected = check, Version = "1", Size = "", Description = "", InstallOrder = id == "B" ? 9 : 0 };
        manager.ReplaceActiveList([Item("A", "工具", true), Item("B", "办公", false), Item("C", "工具", false), Item("D", "工具", true)], ["工具", "办公", "空分类"]);
        Call("RebuildTreeView"); main.Show(); Application.DoEvents();
        var tree = Field<TreeView>("categoryTreeView"); var list = Field<ListView>("softwareListView");
        string Ids() => string.Join("", manager.ActiveSoftwareItems.Select(s => s.Id));
        tree.SelectedNode = tree.Nodes[0].Nodes[2]; Field<Button>("treeMoveUp").PerformClick();
        Assert(manager.ActiveCategories.SequenceEqual(["工具", "空分类", "办公"]) && tree.SelectedNode!.Text == "空分类", "moving an empty category preserves selection and custom category order");
        Field<Button>("treeMoveDown").PerformClick(); Field<Button>("treeMoveDown").PerformClick();
        Assert(manager.ActiveCategories.SequenceEqual(["工具", "办公", "空分类"]) && !Field<Button>("treeMoveDown").Enabled, "category boundaries disable further movement");
        tree.SelectedNode = tree.Nodes[0].Nodes[0].Nodes[1]; Field<Button>("treeMoveUp").PerformClick();
        Assert(Ids() == "CBAD" && ((SoftwareItem)tree.SelectedNode!.Tag!).Id == "C", "tree software movement stays in its category while preserving other categories' slots");
        var category = Field<ComboBox>("softwareCategory"); category.SelectedIndex = 1; Application.DoEvents();
        list.Items[1].Selected = true; list.Items[1].Focused = true; Field<Button>("softwareMoveUp").PerformClick();
        Assert(Ids() == "ABCD" && ((SoftwareItem)list.SelectedItems[0].Tag!).Id == "A", "filtered software movement keeps hidden entries in place and retains the highlighted row");
        list.Items[1].Selected = true; Field<Button>("softwareMoveDown").PerformClick();
        Assert(Ids() == "DBAC" && list.SelectedItems.Count == 2 && !Field<Button>("softwareMoveDown").Enabled, "multiple selected rows move together in stable order and stop at the boundary");
        Assert(manager.ActiveSoftwareItems.Where(s => s.IsSelected).Select(s => s.Id).ToHashSet().SetEquals(["A", "D"]) && manager.ActiveSoftwareItems.Single(s => s.Id == "B").InstallOrder == 9,
            "moving highlights preserves installation checkboxes and explicit installation priorities");
        tree.SelectedNode = tree.Nodes[0];
        Assert(!Field<Button>("treeMoveUp").Enabled && !Field<Button>("treeMoveDown").Enabled, "the all-software root cannot be moved");
        tree.SelectedNode = tree.Nodes[0].Nodes[1]; Field<Button>("treeMoveUp").PerformClick();
        var configuration = (WorkspaceConfiguration)Call("CaptureConfiguration")!;
        var path = Path.Combine(root, "ordered.json"); ConfigurationStore.Save(path, configuration);
        var loaded = ConfigurationStore.Load(path, out _);
        Call("ApplyConfiguration", loaded, path);
        Assert(Ids() == "DBAC" && tree.Nodes[0].Nodes.Cast<TreeNode>().Select(n => n.Text).SequenceEqual(["办公", "工具", "空分类"]),
            "saving and loading the configuration retains both software and category ordering");
        var handler = new SchemeHandler(useDefaultStorage: false);
        handler.SaveScheme("排序方案", manager.ActiveSoftwareItems, manager.ActiveCategories);
        using var wizard = new BundleBuildForm(manager.ActiveSoftwareItems, handler, manager.ActiveCategories);
        var request = (BundleBuildRequest)typeof(BundleBuildForm).GetMethod("Prepare", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(wizard, null)!;
        Assert(request.Software.Select(s => s.Id).SequenceEqual(["D", "B", "A", "C"]) && request.Categories.SequenceEqual(["办公", "工具", "空分类"])
            && handler.GetCategories("排序方案").SequenceEqual(request.Categories), "scheme snapshots and generation requests retain manual ordering");
        main.Size = main.MinimumSize; Application.DoEvents();
        using (var picture = new Bitmap(main.Width, main.Height)) { main.DrawToBitmap(picture, new Rectangle(Point.Empty, main.Size)); picture.Save(Path.Combine(root, "preview-ordering-minimum.png")); }
        main.Size = new(1080, 720); Application.DoEvents();
        using (var picture = new Bitmap(main.Width, main.Height)) { main.DrawToBitmap(picture, new Rectangle(Point.Empty, main.Size)); picture.Save(Path.Combine(root, "preview-ordering.png")); }
    }
}
