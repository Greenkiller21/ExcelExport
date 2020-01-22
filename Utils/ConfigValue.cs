using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExport.Utils
{
    public class ConfigValue
    {
        public string Key { get; set; }
        public string Value { get; set; }

        public ConfigValue(string name, string value)
        {
            Key = name;
            Value = value;
        }
    }
}
