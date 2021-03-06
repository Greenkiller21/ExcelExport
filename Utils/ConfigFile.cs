﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport.Utils
{
    public class ConfigFile
    {
        private string _pathToConfig = "";

        public List<ConfigSection> Sections { get; set; } = new List<ConfigSection>();

        public ConfigFile(string pathToConfig)
        {
            _pathToConfig = pathToConfig;
            EnsureDirectoryCreated(_pathToConfig);
        }

        private void EnsureDirectoryCreated(string path)
        {
            var info = new FileInfo(path);

            if (!Directory.Exists(info.DirectoryName))
                Directory.CreateDirectory(info.DirectoryName);
        }

        public void Load()
        {
            if (!File.Exists(_pathToConfig))
                File.Create(_pathToConfig).Close();

            var file = new StreamReader(_pathToConfig);
            var currentSection = new ConfigSection("default");

            while (!file.EndOfStream)
            {
                var line = file.ReadLine().Trim();

                if (line.Length > 0)
                {
                    if (line.StartsWith("[") && line.EndsWith("]"))
                    {
                        line = line.Replace("[", string.Empty).Replace("]", string.Empty);

                        if (currentSection != null && currentSection.Values.Count > 0)
                            Sections.Add(currentSection);
                        currentSection = new ConfigSection(line);
                    }
                    else if (line.Contains("="))
                    {
                        var content = line.Split('=');
                        var key = content[0].Trim();
                        var value = content[1].Trim();

                        currentSection.Values.Add(new ConfigValue(key, value));
                    }
                }
            }

            if (currentSection != null && currentSection.Values.Count > 0)
                Sections.Add(currentSection);

            file.Close();
        }

        public void Save()
        {
            var file = new StreamWriter(_pathToConfig);

            foreach (var section in Sections)
            {
                if (section.Name != "default" && section.Values.Count > 0)
                    file.WriteLine($"[{section.Name}]");

                foreach (var value in section.Values)
                {
                    file.WriteLine($"{value.Key}={value.Value}");
                }
            }

            file.Close();
        }

        public ConfigSection this[string key]
        {
            get 
            {
                var section = Sections.FirstOrDefault(x => x.Name == key);
                if (section == null)
                {
                    section = new ConfigSection(key);
                    Sections.Add(section);
                }

                return Sections.FirstOrDefault(x => x.Name == key);
            }
            set
            {
                var section = Sections.FirstOrDefault(x => x.Name == value.Name);
                Sections.Remove(section);
                Sections.Add(value);
            }
        }
    }
}
