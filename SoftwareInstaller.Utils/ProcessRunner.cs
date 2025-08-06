using SoftwareInstaller.Models;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SoftwareInstaller.Utils
{
    public static class ProcessRunner
    {
        public static Process? StartProcess(string filePath, string args)
        {
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = filePath,
                    Arguments = args,
                    UseShellExecute = true,
                    Verb = "runas" // 以管理员权限运行
                };
                return Process.Start(startInfo);
            }
            catch (System.Exception)
            {
                return null;
            }
        }

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
                startInfo.UseShellExecute = true;
                startInfo.Verb = "runas";
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

        public static async Task<string> ExecuteWithOutput(string command, string args)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = command,
                Arguments = args,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                Verb = "runas"
            };

            using (Process process = new Process { StartInfo = startInfo })
            {
                process.Start();
                string output = await process.StandardOutput.ReadToEndAsync();
                await process.WaitForExitAsync();
                return output;
            }
        }
    }
}
