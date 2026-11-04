using SoftwareInstaller.Core;

namespace SoftwareInstaller.UI;

public partial class MainForm
{
    private void RepairResourcePaths()
    {
        var configuration = CaptureConfiguration();
        if (ResourceRelocation.MissingCount(configuration) == 0)
        {
            MessageBox.Show(this, "当前列表和已保存方案中的安装文件及关联目录均可找到，无需修复。", "修复路径", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        using var dialog = new ResourceRepairForm(configuration);
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        _softwareManager!.ReplaceActiveList(configuration.Software, configuration.Categories);
        _schemeHandler!.ReplaceSchemes(configuration.Schemes);
        configurationRevision++; // Discard any imports prepared against the previous paths.
        RebuildTreeView();
        MessageBox.Show(this, $"已修复 {dialog.AppliedCount} 条软件记录。请点击「保存配置」保存修改。", "修复完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}
