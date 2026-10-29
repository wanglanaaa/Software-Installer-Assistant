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
        public List<string> ActiveCategories { get; } = new();
        private readonly CategoryStore categoryStore = new();
        private const string ConfigPath = "software_list.json";
        private readonly bool useDefaultStorage;


        public SoftwareManager(bool useDefaultStorage = true)
        {
            this.useDefaultStorage = useDefaultStorage;
            _knownSoftwareItems = useDefaultStorage ? LoadSoftwareItems() : new();
            ActiveSoftwareItems = new List<SoftwareItem>(); // 启动时，活动列表为空
            if (useDefaultStorage) ActiveCategories.AddRange(categoryStore.Load(_knownSoftwareItems.Select(s => s.Category)));
        }

        public void AddCategory(string name)
        {
            name = name.Trim();
            if (name.Length == 0 || ActiveCategories.Contains(name, StringComparer.OrdinalIgnoreCase)) return;
            if (useDefaultStorage) categoryStore.Save(ActiveCategories.Append(name));
            ActiveCategories.Add(name);
        }

        public void RemoveCategory(string name)
        {
            if (useDefaultStorage) categoryStore.Save(ActiveCategories.Where(c => !c.Equals(name, StringComparison.OrdinalIgnoreCase)));
            ActiveCategories.RemoveAll(c => c.Equals(name, StringComparison.OrdinalIgnoreCase));
            ActiveSoftwareItems.RemoveAll(s => s.Category.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// 将一个新软件添加到活动列表，并从知识库中为其填充已知信息
        /// </summary>
        /// <param name="newSoftware">从UI层传入的新软件对象</param>
        public void AddSoftwareToActiveList(SoftwareItem newSoftware)
        {
            // 尝试在知识库中寻找匹配的软件（这里使用软件名称作为唯一标识）
            var knownItem = _knownSoftwareItems.FirstOrDefault(item => item.Name.Equals(newSoftware.Name, StringComparison.OrdinalIgnoreCase));

            if (knownItem != null && string.IsNullOrWhiteSpace(newSoftware.SilentInstallArgs))
            {
                // 如果找到了，就用知识库中的信息（特别是静默参数）来更新新软件对象
                newSoftware.SilentInstallArgs = knownItem.SilentInstallArgs;
            }

            // 将处理过的新软件添加到活动列表中
            var existing = ActiveSoftwareItems.FindIndex(s => s.Id == newSoftware.Id);
            if (existing >= 0) ActiveSoftwareItems[existing] = newSoftware;
            else ActiveSoftwareItems.Add(newSoftware);
            AddCategory(newSoftware.Category);
        }

        /// <summary>
        /// 将当前活动列表的更改合并回知识库，并完整保存到JSON文件
        /// </summary>
        public void SaveSoftwareList()
        {
            if (!useDefaultStorage) return;
            try
            {
                // Save preparation work too; verification remains an explicit separate state.
                var valuableItems = ActiveSoftwareItems;

                foreach (var valuableItem in valuableItems)
                {
                    var knownItem = _knownSoftwareItems.FirstOrDefault(item => item.Id == valuableItem.Id);
                    if (knownItem != null)
                    {
                        // 如果知识库中已存在，则用最新的、有价值的信息更新它
                        _knownSoftwareItems[_knownSoftwareItems.IndexOf(knownItem)] = valuableItem.Clone();
                    }
                    else
                    {
                        // 如果是全新的、且有价值的软件，则添加到知识库
                        _knownSoftwareItems.Add(valuableItem.Clone());
                    }
                }

                // 将完整的、高质量的知识库序列化并写入文件
                string jsonString = JsonSerializer.Serialize(_knownSoftwareItems, new JsonSerializerOptions { WriteIndented = true });
                UserDataPaths.Write(ConfigPath, jsonString);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving software list: {ex.Message}");
            }
        }

        public void ReplaceActiveList(IEnumerable<SoftwareItem> software, IEnumerable<string> categories)
        {
            var copies = software.Select(s => s.Clone()).ToList();
            var names = CategoryNames.Normalize(categories.Concat(copies.Select(s => s.Category)));
            ActiveSoftwareItems.Clear();
            ActiveSoftwareItems.AddRange(copies);
            ActiveCategories.Clear();
            ActiveCategories.AddRange(names);
        }

        private List<SoftwareItem> LoadSoftwareItems()
        {
            var path = UserDataPaths.ReadPath(ConfigPath);
            if (!File.Exists(path))
            {
                return new List<SoftwareItem>();
            }

            try
            {
                string jsonString = File.ReadAllText(path);
                return JsonSerializer.Deserialize<List<SoftwareItem>>(jsonString) ?? new List<SoftwareItem>();
            }
            catch (Exception ex)
            {
                throw new Exception($"加载软件列表失败。\n请检查 software_list.json 文件的格式是否正确。", ex);
            }
        }
    }
}
