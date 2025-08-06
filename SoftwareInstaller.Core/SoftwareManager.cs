using SoftwareInstaller.Models;
using SoftwareInstaller.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;

namespace SoftwareInstaller.Core
{
    public class SoftwareManager
    {
        private List<SoftwareItem> _knownSoftwareItems; // 内部知识库
        public List<SoftwareItem> ActiveSoftwareItems { get; private set; } // 当前UI操作的列表
        private const string ConfigPath = "software_list.json";

        private static readonly List<string> SilentArgsList = new List<string>
        {
            "/S", "/VERYSILENT", "/SILENT", "/q", "/qn", "/s", "/quiet"
        };

        public SoftwareManager()
        {
            _knownSoftwareItems = LoadSoftwareItems();
            ActiveSoftwareItems = new List<SoftwareItem>(); // 启动时，活动列表为空
        }

        /// <summary>
        /// 将一个新软件添加到活动列表，并从知识库中为其填充已知信息
        /// </summary>
        /// <param name="newSoftware">从UI层传入的新软件对象</param>
        public void AddSoftwareToActiveList(SoftwareItem newSoftware)
        {
            // 尝试在知识库中寻找匹配的软件（这里使用软件名称作为唯一标识）
            var knownItem = _knownSoftwareItems.FirstOrDefault(item => item.Name.Equals(newSoftware.Name, StringComparison.OrdinalIgnoreCase));

            if (knownItem != null)
            {
                // 如果找到了，就用知识库中的信息（特别是静默参数）来更新新软件对象
                newSoftware.SilentInstallArgs = knownItem.SilentInstallArgs;
                newSoftware.Category = knownItem.Category; // 也可以同步分类等其他信息
            }

            // 将处理过的新软件添加到活动列表中
            ActiveSoftwareItems.Add(newSoftware);
        }

        public async Task<(bool success, string? usedArgs)> InstallSoftwareIntelligently(SoftwareItem item, Func<Task<bool>> installCompletedCallback)
        {
            if (string.IsNullOrEmpty(item.FilePath) || !File.Exists(item.FilePath))
            {
                throw new FileNotFoundException("安装文件未找到。", item.FilePath);
            }

            if (!string.IsNullOrEmpty(item.SilentInstallArgs))
            {
                if (await TryInstall(item.FilePath, item.SilentInstallArgs) && await installCompletedCallback())
                {
                    return (true, item.SilentInstallArgs);
                }
            }

            foreach (var args in SilentArgsList)
            {
                if (await TryInstall(item.FilePath, args) && await installCompletedCallback())
                {
                    item.SilentInstallArgs = args; // 学习到了新参数
                    SaveSoftwareList(); // 立即保存，确保知识不丢失
                    return (true, args);
                }
            }

            return (false, null);
        }

        private async Task<bool> TryInstall(string filePath, string args)
        {
            try
            {
                using var process = ProcessRunner.StartProcess(filePath, args);
                if (process == null) return false;

                await Task.Delay(2000);

                if (WindowDetector.HasVisibleWindow(process.Id))
                {
                    process.Kill();
                    return false;
                }

                var cts = new CancellationTokenSource(TimeSpan.FromMinutes(5));
                await process.WaitForExitAsync(cts.Token);

                return process.ExitCode == 0;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 将当前活动列表的更改合并回知识库，并完整保存到JSON文件
        /// </summary>
        public void SaveSoftwareList()
        {
            try
            {
                // 过滤出那些已经成功学习到静默参数的活动项
                var valuableItems = ActiveSoftwareItems.Where(item => !string.IsNullOrEmpty(item.SilentInstallArgs));

                foreach (var valuableItem in valuableItems)
                {
                    var knownItem = _knownSoftwareItems.FirstOrDefault(item => item.Name.Equals(valuableItem.Name, StringComparison.OrdinalIgnoreCase));
                    if (knownItem != null)
                    {
                        // 如果知识库中已存在，则用最新的、有价值的信息更新它
                        knownItem.Version = valuableItem.Version;
                        knownItem.Description = valuableItem.Description;
                        knownItem.Category = valuableItem.Category;
                        knownItem.FilePath = valuableItem.FilePath; // 更新路径以反映最新选择
                        knownItem.SilentInstallArgs = valuableItem.SilentInstallArgs; // 更新或确认静默参数
                    }
                    else
                    {
                        // 如果是全新的、且有价值的软件，则添加到知识库
                        _knownSoftwareItems.Add(valuableItem);
                    }
                }

                // 将完整的、高质量的知识库序列化并写入文件
                string jsonString = JsonSerializer.Serialize(_knownSoftwareItems, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigPath, jsonString);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving software list: {ex.Message}");
            }
        }

        private List<SoftwareItem> LoadSoftwareItems()
        {
            if (!File.Exists(ConfigPath))
            {
                return new List<SoftwareItem>();
            }

            try
            {
                string jsonString = File.ReadAllText(ConfigPath);
                return JsonSerializer.Deserialize<List<SoftwareItem>>(jsonString) ?? new List<SoftwareItem>();
            }
            catch (Exception ex)
            {
                throw new Exception($"加载软件列表失败。\n请检查 software_list.json 文件的格式是否正确。", ex);
            }
        }
    }
}
