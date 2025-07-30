using SoftwareInstaller.Models;
using System.Diagnostics;
using System.Windows.Forms;

namespace SoftwareInstaller.Utils
{
    public static class ProcessRunner
    {
        public static void ExecuteInstallation(SoftwareItem item, bool isAuto)
        {
            if (string.IsNullOrEmpty(item.FilePath))
            {
                MessageBox.Show($"软件 '{item.Name}' 未找到安装文件路径。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.UseShellExecute = true; // 必须设置为 true 才能使用 Verb
                startInfo.Verb = "runas";         // "runas" 表示请求管理员权限
                startInfo.FileName = item.FilePath;
                if (isAuto)
                {
                    if (string.IsNullOrEmpty(item.SilentInstallArgs))
                    {
                        MessageBox.Show($"软件 '{item.Name}' 未提供静默安装参数，无法自动安装。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    startInfo.Arguments = item.SilentInstallArgs;
                }
                Process.Start(startInfo);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"无法启动安装程序 '{item.Name}'.\n错误: {ex.Message}", "安装失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
