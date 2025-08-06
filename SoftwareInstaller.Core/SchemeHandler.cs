using SoftwareInstaller.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace SoftwareInstaller.Core
{
    public class SchemeHandler
    {
        private static readonly string SchemesFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "schemes.json");
        private readonly Dictionary<string, List<SoftwareItem>> _schemes;

        public SchemeHandler()
        {
            _schemes = LoadSchemesFromFile();
        }

        public IEnumerable<string> GetSchemeNames()
        {
            return _schemes.Keys.OrderBy(k => k);
        }

        public bool SchemeExists(string name)
        {
            return _schemes.ContainsKey(name);
        }

        public List<SoftwareItem>? GetScheme(string name)
        {
            _schemes.TryGetValue(name, out var items);
            return items;
        }

        public void SaveScheme(string name, List<SoftwareItem> items)
        {
            // 克隆项目，确保保存的方案是一个快照
            _schemes[name] = new List<SoftwareItem>(items.Select(item => item.Clone()));
            SaveSchemesToFile();
        }

        public void DeleteScheme(string name)
        {
            if (_schemes.Remove(name))
            {
                SaveSchemesToFile();
            }
        }

        private Dictionary<string, List<SoftwareItem>> LoadSchemesFromFile()
        {
            if (!File.Exists(SchemesFilePath))
                return new Dictionary<string, List<SoftwareItem>>();
            try
            {
                var json = File.ReadAllText(SchemesFilePath);
                return JsonSerializer.Deserialize<Dictionary<string, List<SoftwareItem>>>(json) ?? new Dictionary<string, List<SoftwareItem>>();
            }
            catch
            {
                return new Dictionary<string, List<SoftwareItem>>();
            }
        }

        private void SaveSchemesToFile()
        {
            var json = JsonSerializer.Serialize(_schemes, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(SchemesFilePath, json);
        }
    }
}
