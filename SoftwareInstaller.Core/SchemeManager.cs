using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using SoftwareInstaller.Models;
using SoftwareInstaller.Utils;

namespace SoftwareInstaller.Core
{
    public static class SchemeManager
    {
        private static readonly string AppDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        private static readonly string AppFolderPath = Path.Combine(AppDataPath, "SoftwareInstaller");
        private static readonly string SchemeFilePath = Path.Combine(AppFolderPath, "schemes.json");

        public static Dictionary<string, List<SoftwareItem>> LoadSchemes()
        {
            try
            {
                if (File.Exists(SchemeFilePath))
                {
                    string jsonString = File.ReadAllText(SchemeFilePath);
                    return JsonSerializer.Deserialize<Dictionary<string, List<SoftwareItem>>>(jsonString) ?? new Dictionary<string, List<SoftwareItem>>();
                }
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowError($"加载方案文件失败.\n错误: {ex.Message}");
            }
            return new Dictionary<string, List<SoftwareItem>>();
        }

        public static void SaveSchemes(Dictionary<string, List<SoftwareItem>> schemes)
        {
            try
            {
                Directory.CreateDirectory(AppFolderPath);
                string jsonString = JsonSerializer.Serialize(schemes, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(SchemeFilePath, jsonString);
            }
            catch (Exception ex)
            {
                ErrorHandler.ShowError($"保存方案文件失败.\n错误: {ex.Message}");
            }
        }
    }
}
