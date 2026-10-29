using SoftwareInstaller.Models;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SoftwareInstaller.Utils
{
    public static class ProcessRunner
    {
        public static Process? StartProcess(SoftwareItem item, string args, bool silent = true)
        {
            if (item.DirectoryMode != InstallDirectoryMode.Automatic || InstallerInspector.Inspect(item.FilePath).DirectoryMode != InstallDirectoryMode.Automatic)
                InstallDirectoryOptions.CheckTargetDrive(item.DefaultInstallDirectory);
            var arguments = InstallDirectoryOptions.BuildArguments(item, args);
            bool isMsi = string.Equals(System.IO.Path.GetExtension(item.FilePath), ".msi", System.StringComparison.OrdinalIgnoreCase);
            return StartProcess(isMsi ? System.IO.Path.Combine(System.Environment.SystemDirectory, "msiexec.exe") : item.FilePath!,
                isMsi ? $"/i {InstallDirectoryOptions.Quote(item.FilePath!)} {arguments} {(silent ? "/qn /norestart REBOOT=ReallySuppress" : "")}" : arguments);
        }
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
                if (isAuto && string.IsNullOrEmpty(item.SilentInstallArgs) && !string.Equals(System.IO.Path.GetExtension(item.FilePath), ".msi", System.StringComparison.OrdinalIgnoreCase))
                    throw new System.InvalidOperationException("未提供静默参数，无法自动安装。");
                using var process = StartProcess(item, isAuto ? item.SilentInstallArgs ?? "" : "", isAuto);
                if (process == null) throw new System.InvalidOperationException("安装器未启动，可能取消了管理员授权。");
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
