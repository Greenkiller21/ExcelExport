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
        private Spire.Xls.Workbook excelFileSpire;
        private Application excelFileInterop;
        private Workbook excelBook;

        public List<ExcelSheet> ExcelSheets { get; private set; }

        public ExcelFile(string filePath)
        {
            AppDomain.CurrentDomain.ProcessExit += FileChanged;
            Views.CustomComponents.Shared.FileChangedEvent += FileChanged;

            ExcelSheets = new List<ExcelSheet>();

            //For image
            excelFileSpire = new Spire.Xls.Workbook();
            excelFileSpire.LoadFromFile(filePath);

            //For PDF
            excelFileInterop = new Application();
            excelBook = excelFileInterop.Workbooks.Open(filePath);

            for(int idx = 0; idx < excelFileSpire.Worksheets.Count; idx++)
            {
                ExcelSheets.Add(new ExcelSheet(excelFileSpire.Worksheets[idx], excelBook.Sheets[idx+1]));
            }
        }

        private void FileChanged(object sender, EventArgs e)
        {
            Close();

            AppDomain.CurrentDomain.ProcessExit -= FileChanged;
            Views.CustomComponents.Shared.FileChangedEvent -= FileChanged;
        }

        public void Close()
        {
            excelBook.Close();
        }
    }
}
