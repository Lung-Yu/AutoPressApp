using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using AutoPressApp.Models;
using AutoPressApp.Services;

namespace AutoPressApp
{
    public static class WorkflowStorage
    {
        public class SavedWorkflowInfo
        {
            public string Name { get; set; } = string.Empty;
            public string FilePath { get; set; } = string.Empty;
            public DateTime LastWriteTime { get; set; }
            public override string ToString() => Name;
        }

        private static readonly object _lock = new();
        public static string RootDir => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Workflows");

        public static void EnsureDir()
        {
            try { Directory.CreateDirectory(RootDir); } catch { }
        }

        public static string SanitizeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "workflow";
            var invalid = Path.GetInvalidFileNameChars();
            var parts = name.Split(invalid, StringSplitOptions.RemoveEmptyEntries);
            var safe = string.Join("_", parts).Trim();
            if (string.IsNullOrWhiteSpace(safe)) safe = "workflow";
            return safe;
        }

        public static string GetFilePath(string name)
        {
            EnsureDir();
            var safe = SanitizeName(name);
            return Path.Combine(RootDir, safe + ".json");
        }

        public static void Save(Workflow wf)
        {
            EnsureDir();
            var json = WorkflowRunner.SaveToJson(wf);
            var path = GetFilePath(wf.Name);
            lock (_lock)
            {
                File.WriteAllText(path, json);
            }
        }

        public static Workflow? Load(string name)
        {
            var path = GetFilePath(name);
            if (!File.Exists(path)) return null;
            try
            {
                var json = File.ReadAllText(path);
                return WorkflowRunner.LoadFromJson(json);
            }
            catch { return null; }
        }

        public static List<SavedWorkflowInfo> List()
        {
            EnsureDir();
            var list = new List<SavedWorkflowInfo>();
            try
            {
                foreach (var file in Directory.GetFiles(RootDir, "*.json"))
                {
                    try
                    {
                        var json = File.ReadAllText(file);
                        var wf = WorkflowRunner.LoadFromJson(json);
                        var info = new SavedWorkflowInfo
                        {
                            Name = string.IsNullOrWhiteSpace(wf.Name) ? Path.GetFileNameWithoutExtension(file) : wf.Name,
                            FilePath = file,
                            LastWriteTime = File.GetLastWriteTime(file)
                        };
                        list.Add(info);
                    }
                    catch { }
                }
            }
            catch { }
            return list.OrderByDescending(x => x.LastWriteTime).ToList();
        }

        public static bool Delete(string name)
        {
            try
            {
                var path = GetFilePath(name);
                if (File.Exists(path))
                {
                    File.Delete(path);
                    return true;
                }
            }
            catch { }
            return false;
        }
    }
}
