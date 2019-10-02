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

        public List<ExcelSheet> ExcelSheets { get; private set; }

        public ExcelFile(string filePath)
        {
            ExcelSheets = new List<ExcelSheet>();

            //For image
            Spire.Xls.Workbook excelFileSpire = new Spire.Xls.Workbook();
            excelFileSpire.LoadFromFile(filePath);

            //For PDF
            Application excelFileInterop = new Application();
            Workbook excelBook = excelFileInterop.Workbooks.Open(filePath);

            for(int idx = 0; idx < excelFileSpire.Worksheets.Count; idx++)
            {
                ExcelSheets.Add(new ExcelSheet(excelFileSpire.Worksheets[idx], excelBook.Sheets[idx+1]));
            }
        }
    }
}
