using System;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace ExcelExport
{
    public class Program
    {
        [STAThreadAttribute]
        public static void Main()
        {
            CosturaUtility.Initialize();

            App.Main();
        }
    }
}
