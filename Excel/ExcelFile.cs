using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelExport.Excel
{
    class ExcelFile
    {

        public List<ExcelSheet> excelSheets { get; private set; }

        public ExcelFile(string filePath)
        {
            Application excelFile = new Application();
            

            foreach()
            {
                excelSheets.Add(new ExcelSheet());
            }
        }
    }
}
