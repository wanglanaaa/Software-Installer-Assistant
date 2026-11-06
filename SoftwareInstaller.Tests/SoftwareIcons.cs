using System.Diagnostics;
using System.Reflection;
using SoftwareInstaller.Core;
using SoftwareInstaller.Models;
using SoftwareInstaller.Packaging;
using SoftwareInstaller.Runner;
using SoftwareInstaller.UI;
using SoftwareInstaller.Utils;

namespace SoftwareInstaller.Tests;

internal static partial class Program
{
    private static void TestSoftwareIcons()
    {
        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException);
        Application.EnableVisualStyles();
        T Field<T>(object form, string name) => (T)form.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(form)!;
        void Preview(Form form, string name)
        {
            form.ShowInTaskbar = false; form.Opacity = 0; form.Show(); Application.DoEvents();
            using var bitmap = new Bitmap(form.Width, form.Height);
            form.DrawToBitmap(bitmap, new Rectangle(Point.Empty, form.Size)); bitmap.Save(Path.Combine(root, name + ".png"));
        }
        var source = Path.Combine(root, "注册表编辑器.exe");
        File.Copy(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "regedit.exe"), source);
        var msi = Path.Combine(root, "示例安装包.msi"); File.WriteAllText(msi, "icon fallback fixture");
        var items = new List<SoftwareItem>
        {
            new() { Name = "注册表编辑器", FilePath = source, Category = "工具", IsSelected = true, Version = "1", Size = "", Description = "安装源自带图标" },
            new() { Name = "示例安装包", FilePath = msi, Category = "工具", Version = "1", Size = "", Description = "默认文件图标" }
        };
        using var fallback = SoftwareIcons.FromFile(Path.Combine(root, "missing.exe"));
        using var invalid = SoftwareIcons.FromPng([1, 2, 3]);
        Assert(fallback.Width > 0 && invalid.Width > 0, "missing files and corrupt icon data have a usable fallback");
        using var main = new MainForm { Size = new(1080, 720) };
        var manager = Field<SoftwareManager>(main, "_softwareManager");
        manager.ReplaceActiveList(items, ["工具"]);
        typeof(MainForm).GetMethod("RebuildTreeView", BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(main, null);
        Preview(main, "preview-software-icons");
        var list = Field<ListView>(main, "softwareListView");
        var images = Field<ImageList>(main, "softwareIcons");
        Assert(list.Items.Count == 2 && list.Items.Cast<ListViewItem>().All(r => r.ImageIndex > 0 && r.ImageIndex < images.Images.Count),
            "each main-list software row has its own name-column icon");
        var tree = Field<TreeView>(main, "categoryTreeView");
        Assert(tree.Nodes[0].Nodes[0].Nodes.Cast<TreeNode>().All(n => n.ImageIndex > 0 && n.SelectedImageIndex == n.ImageIndex),
            "sidebar software names retain their icons when selected");
        Field<TextBox>(main, "softwareSearch").Text = items[0].Name;
        Assert(list.Items.Count == 1 && list.Items[0].ImageIndex > 0 && list.Items[0].Checked, "filtering retains software icons and checkboxes");
        Field<TextBox>(main, "softwareSearch").Clear();
        list.Items[0].Selected = true; list.Focus(); Application.DoEvents();
        Preview(main, "preview-software-icons-selected");

        var output = Path.Combine(root, "图标合集.exe");
        Task.Run(() => BundleBuilder.BuildAsync(new() { Name = "图标合集", RunnerPath = Environment.ProcessPath!, OutputPath = output,
            Software = items, DefaultScheme = "全部", Schemes = new() { ["全部"] = items.Select(s => s.Id).ToList() } })).GetAwaiter().GetResult();
        using (var package = Task.Run(() => BundleReader.OpenAsync(output)).GetAwaiter().GetResult())
        {
            Assert(package.Manifest.Software.All(s => s.IconPng is { Length: > 0 }), "generated package embeds each software icon");
            Assert(package.Manifest.Software[0].IconPng!.SequenceEqual(SoftwareIcons.ReadPng(source)), "packaged icon matches the original installer icon");
        }
        File.Move(source, source + ".moved"); File.Move(msi, msi + ".moved");
        using var runnerForm = new InstallForm(output);
        runnerForm.ShowInTaskbar = false; runnerForm.Opacity = 0; runnerForm.Show();
        var install = Field<Button>(runnerForm, "install"); var watch = Stopwatch.StartNew();
        while (!install.Enabled && watch.Elapsed < TimeSpan.FromSeconds(10)) { Application.DoEvents(); Thread.Sleep(10); }
        var runnerList = Field<ListView>(runnerForm, "list");
        Assert(install.Enabled && runnerList.Items.Count == 2 && runnerList.Items.Cast<ListViewItem>().All(r => r.ImageIndex >= 0),
            "exported interface loads embedded icons after original installers have moved");
        Preview(runnerForm, "preview-runner-software-icons");
        runnerList.Items[0].Checked = false;
        Assert(runnerList.CheckedItems.Count == 1, "runner selection still works with software icons");
    }
}
