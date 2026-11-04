using System.Diagnostics;
using System.Reflection;
using SoftwareInstaller.Core;
using SoftwareInstaller.Models;
using SoftwareInstaller.Packaging;
using SoftwareInstaller.UI;

namespace SoftwareInstaller.Tests;

internal static partial class Program
{
    private static void TestBundleOverview()
    {
        var folder = Path.Combine(root, "关联资源"); Directory.CreateDirectory(Path.Combine(folder, "data"));
        var exe = Path.Combine(folder, "setup.exe"); File.WriteAllBytes(exe, [77, 90, 0]);
        File.WriteAllBytes(Path.Combine(folder, "extra.bin"), new byte[17]); File.WriteAllBytes(Path.Combine(folder, "data", "配置.bin"), new byte[13]);
        var msi = Path.Combine(root, "联网安装.msi"); File.WriteAllBytes(msi, [77, 90, 1, 2]);
        SoftwareItem Item(string id, string name, string file) => new() { Id = id, Name = name, FilePath = file, IsSelected = true, Category = "工具", Version = "1", Description = "", Size = "不可信的旧大小" };
        var interactive = Item("a", "办公助手", exe); interactive.ResourceDirectory = folder; interactive.DefaultInstallDirectory = "D:\\Apps";
        var online = Item("b", "联网工具", msi); online.RequiresNetwork = true;
        var shared = Item("c", "共用资源工具", exe); shared.ResourceDirectory = folder; shared.SilentInstallArgs = "--silent";
        var request = new BundleBuildRequest { Software = [interactive, online, shared], Categories = ["工具", "空分类"], RunnerPath = runner,
            Schemes = new() { ["全部"] = ["a", "b", "c"], ["常用"] = ["a"] } };
        var descriptions = request.Software.Select(BundleBuilder.Describe).ToList();
        var summary = BundleOverview.Create(request, descriptions);
        Assert(summary.SoftwareCount == 3 && summary.SchemeCount == 2 && summary.CategoryCount == 2, "overview counts the prepared software, schemes and empty categories");
        Assert(summary.SourceBytes == 70 && summary.FileCount == 7, "resource sizes come from actual nested files and count duplicate references needed for extraction");
        Assert(summary.BuildFreeBytes == 140 + summary.RunnerBytes + BundleOverview.SpaceHeadroom && summary.TargetTempBytes == 70 + BundleOverview.SpaceHeadroom
            && summary.OutputReferenceBytes == summary.RunnerBytes + 70 + BundleReader.FooterSize, "space estimates match builder scratch-space rules and include runner overhead");
        Assert(summary.NetworkSoftware.SequenceEqual(["联网工具"]) && summary.InteractiveSoftware.SequenceEqual(["办公助手"]) && summary.DirectorySelectionSoftware.SequenceEqual(["办公助手"]),
            "overview distinguishes network, interactive installation and manual directory selection");
        using (var locked = new FileStream(exe, FileMode.Open, FileAccess.Read, FileShare.None))
            Assert(BundleOverview.Create(request, descriptions).SourceBytes == 70, "overview reuses installer descriptions and needs no content hashing or reading of resource data");
        Assert(BundleOverview.FormatBytes(1024) == "1 KiB" && BundleOverview.FormatBytes(1024L * 1024 * 1024) == "1 GiB", "binary size labels use the correct units");
        var handler = new SchemeHandler(useDefaultStorage: false); handler.SaveScheme("附加", [shared], ["工具", "空分类"]);
        Application.EnableVisualStyles();
        using var wizard = new BundleBuildForm([interactive, online], handler, ["工具", "空分类"]) { Opacity = 0, ShowInTaskbar = false };
        T Field<T>(string name) => (T)typeof(BundleBuildForm).GetField(name, BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(wizard)!;
        object? Call(string name, params object[] arguments) => typeof(BundleBuildForm).GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic)!.Invoke(wizard, arguments);
        void Check()
        {
            var task = (Task)Call("CheckAsync")!; var elapsed = Stopwatch.StartNew();
            while (!task.IsCompleted && elapsed.Elapsed < TimeSpan.FromSeconds(15)) { Application.DoEvents(); Thread.Sleep(5); }
            if (!task.IsCompleted) throw new Exception("Overview check timed out.");
            task.GetAwaiter().GetResult();
        }
        wizard.Show(); Application.DoEvents(); Field<CheckedListBox>("schemes").SetItemChecked(0, true); Call("SetPage", 1); Check();
        Assert(Field<bool>("checkedReady") && Field<BundleOverview>("overview").SoftwareCount == 3 && Field<Label>("checkOverview").Text.Contains("需要联网 1 款"),
            "wizard overview reflects attached schemes and current installation requirements");
        using (var picture = new Bitmap(wizard.Width, wizard.Height)) { wizard.DrawToBitmap(picture, new Rectangle(Point.Empty, wizard.Size)); picture.Save(Path.Combine(root, "preview-overview-checks.png")); }
        Field<Button>("next").PerformClick(); Field<TextBox>("output").Text = Path.Combine(root, "preview.exe"); Application.DoEvents();
        Assert(Field<TextBox>("overviewText").Text.Contains("实际大小由压缩和内容去重决定") && Field<TextBox>("overviewText").Text.Contains("办公助手") && Field<Label>("outputSpace").Text.Contains("可用"),
            "generation page explains estimated size, names affected software and checks the selected output disk");
        using (var picture = new Bitmap(wizard.Width, wizard.Height)) { wizard.DrawToBitmap(picture, new Rectangle(Point.Empty, wizard.Size)); picture.Save(Path.Combine(root, "preview-overview-generate.png")); }
        wizard.Size = wizard.MinimumSize; Application.DoEvents();
        using (var picture = new Bitmap(wizard.Width, wizard.Height)) { wizard.DrawToBitmap(picture, new Rectangle(Point.Empty, wizard.Size)); picture.Save(Path.Combine(root, "preview-overview-minimum.png")); }
        Call("SetPage", 0);
        Assert(Field<BundleOverview?>("overview") == null && Field<TextBox>("overviewText").Text.Length == 0, "returning to content selection clears stale size and requirement summaries");
        // A missing source must not leave a previous successful overview available.
        File.Move(msi, Path.Combine(root, "moved.msi")); Call("SetPage", 1); Check();
        Assert(!Field<bool>("checkedReady") && Field<BundleOverview?>("overview") == null && !Field<Label>("checkOverview").Text.Contains("本次包含"),
            "failed rechecks discard previously successful overview results");
    }
}
