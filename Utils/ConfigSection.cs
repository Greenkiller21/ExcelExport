using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport.Utils
{
    public class ConfigSection
    {
        public string Name { get; set; }

        public List<ConfigValue> Values { get; set; } = new List<ConfigValue>();

        public ConfigSection(string name)
        {
            Name = name;
        }

        public ConfigValue this[string key]
        {
            get
            {
                var value = Values.FirstOrDefault(x => x.Key == key);
                if (value == null)
                {
                    value = new ConfigValue(key, "");
                    Values.Add(value);
                }

                return Values.FirstOrDefault(x => x.Key == key);
            }
            set
            {
                var keyvalue = Values.FirstOrDefault(x => x.Key == value.Key);
                Values.Remove(keyvalue);
                Values.Add(value);
            }
        }
    }
}
