using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelExport.Utils;
using Microsoft.Office.Interop.Excel;
using MvvmCross;
using MvvmCross.ViewModels;

namespace ExcelExport.Excel
{
    public class ExcelFile
    {
        private Spire.Xls.Workbook excelFileSpire;
        private Application excelFileInterop;
        private Workbook excelBook;

        public MvxObservableCollection<ExcelSheet> ExcelSheets { get; private set; }
        public System.Action RefreshAction { get; set; } = () => { };

        public ExcelFile(string filePath)
        {
            ExcelSheets = new MvxObservableCollection<ExcelSheet>();
            
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

        public void Close()
        {
            try
            {
                excelBook.Close();
                excelFileInterop.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelFileInterop);

                var config = Mvx.IoCProvider.GetSingleton<ConfigFile>();
                config.Save();
            }
            catch { }
        }
    }
}
