using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Microsoft.Office.Interop.Excel;

namespace ExcelExport.Excel
{
    class ExcelSheet
    {
        private Spire.Xls.Worksheet excelSheetSpire;
        private Worksheet excelSheetInterop;

        public ExcelSheet(Spire.Xls.Worksheet excelSheetSpire, Worksheet excelSheetInterop)
        {
            this.excelSheetSpire = excelSheetSpire;
            this.excelSheetInterop = excelSheetInterop;
        }

        public Image Previsualisation()
        {
            Image preview;

            using (MemoryStream memoryStream = new MemoryStream())
            {
                excelSheetSpire.ToEMFStream(memoryStream, 1, 1, excelSheetSpire.LastRow, excelSheetSpire.LastColumn);
                preview = Image.FromStream(memoryStream);
            }

            return preview;
        }

        public void ExportToPDF(string filePath, string fileName)
        {
            excelSheetInterop.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, filePath + fileName, XlFixedFormatQuality.xlQualityStandard, true, true, 1, 10, false);
        }
    }
}
