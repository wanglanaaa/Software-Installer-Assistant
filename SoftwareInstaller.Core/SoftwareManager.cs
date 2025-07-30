using SoftwareInstaller.Models;
using SoftwareInstaller.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace SoftwareInstaller.Core
{
    public class SoftwareManager
    {
        public List<SoftwareItem> SoftwareItems { get; private set; }

        public SoftwareManager()
        {
            SoftwareItems = LoadSoftwareItems();
        }

        private List<SoftwareItem> LoadSoftwareItems()
        {
            string configPath = "software_list.json";
            if (!File.Exists(configPath))
            {
                var defaultItems = new List<SoftwareItem>();
                File.WriteAllText(configPath, JsonSerializer.Serialize(defaultItems, new JsonSerializerOptions { WriteIndented = true }));
                return defaultItems;
            }

            try
            {
                string jsonString = File.ReadAllText(configPath);
                return JsonSerializer.Deserialize<List<SoftwareItem>>(jsonString) ?? new List<SoftwareItem>();
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowError($"加载软件列表失败。\n错误: {ex.Message}");
                return new List<SoftwareItem>();
            }
        }
    }
}
