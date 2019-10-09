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
        const int CROP_TOP = 35;
        const int CROP_BOTTOM = 35;
        const int CROP_LEFT = 15;
        const int CROP_RIGHT = 60;

        private Spire.Xls.Worksheet excelSheetSpire;
        private Worksheet excelSheetInterop;

        public string SheetName
        {
            get { return excelSheetInterop.Name; }
        }

        public ExcelSheet(Spire.Xls.Worksheet excelSheetSpire, Worksheet excelSheetInterop)
        {
            this.excelSheetSpire = excelSheetSpire;
            this.excelSheetInterop = excelSheetInterop;
        }

        public Bitmap Preview()
        {
            Bitmap bitmapPreview;

            using (MemoryStream memoryStream = new MemoryStream())
            {
                excelSheetSpire.ToEMFStream(memoryStream, 1, 1, excelSheetSpire.LastRow, excelSheetSpire.LastColumn);
                Image preview = Image.FromStream(memoryStream);

                bitmapPreview = CropImage(preview, CROP_TOP, CROP_BOTTOM, CROP_LEFT, CROP_RIGHT);
            }

            return bitmapPreview;
        }

        public void ExportToPDF(string filePath)
        {
            excelSheetInterop.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, filePath, XlFixedFormatQuality.xlQualityStandard, true, true, 1, 10, false);
        }

        private Bitmap CropImage(Image originalImage, int cropTop, int cropBottom, int cropLeft, int cropRight)
        {
            Bitmap cropImage = new Bitmap(originalImage);
            Bitmap bmpCrop = cropImage.Clone(new System.Drawing.Rectangle(cropLeft, cropTop, cropImage.Width - cropLeft - cropRight, cropImage.Height - cropTop - cropBottom), cropImage.PixelFormat);

            return bmpCrop;
        }
    }
}
