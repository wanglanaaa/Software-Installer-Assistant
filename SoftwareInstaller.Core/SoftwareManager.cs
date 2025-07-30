using SoftwareInstaller.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace SoftwareInstaller.Core
{
    public class SoftwareManager
    {
        public List<SoftwareItem> SoftwareItems { get; private set; }
        private const string ConfigPath = "software_list.json";

        public SoftwareManager(bool empty = false)
        {
            if (empty)
            {
                SoftwareItems = new List<SoftwareItem>();
            }
            else
            {
                SoftwareItems = LoadSoftwareItems();
            }
        }

        public void SaveSoftwareList()
        {
            try
            {
                string jsonString = JsonSerializer.Serialize(SoftwareItems, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigPath, jsonString);
            }
            catch (Exception ex)
            {
                // In a real app, you might want to show this to the user.
                // For now, we just write to console to avoid circular dependency on UI.
                Console.WriteLine($"Error saving software list: {ex.Message}");
            }
        }

        private List<SoftwareItem> LoadSoftwareItems()
        {
            if (!File.Exists(ConfigPath))
            {
                var defaultItems = new List<SoftwareItem>();
                File.WriteAllText(ConfigPath, JsonSerializer.Serialize(defaultItems, new JsonSerializerOptions { WriteIndented = true }));
                return defaultItems;
            }

            try
            {
                string jsonString = File.ReadAllText(ConfigPath);
                return JsonSerializer.Deserialize<List<SoftwareItem>>(jsonString) ?? new List<SoftwareItem>();
            }
            catch (Exception ex)
            {
                // Let the caller (UI layer) handle this exception
                throw new Exception($"加载软件列表失败。\n请检查 software_list.json 文件的格式是否正确。", ex);
            }
        }
    }
}
