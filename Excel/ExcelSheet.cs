using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using System.Windows.Media.Imaging;
using System.Drawing.Imaging;

namespace ExcelExport.Excel
{
    public class ExcelSheet
    {
        const int CROP_TOP = 72;
        const int CROP_BOTTOM = 35;
        const int CROP_LEFT = 68;
        const int CROP_RIGHT = 60;
        const int MIN_IMAGE_SIZE = 200;

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
        
        public BitmapImage Preview()
        {
            Bitmap bitmapPreview;

            using (MemoryStream memoryStream = new MemoryStream())
            {
                try
                {
                    excelSheetSpire.ToEMFStream(memoryStream, 1, 1, excelSheetSpire.LastRow, excelSheetSpire.LastColumn);
                }
                catch
                {
                    return null;
                }

                Image preview = Image.FromStream(memoryStream);

                bitmapPreview = CropImage(preview, CROP_TOP, CROP_BOTTOM, CROP_LEFT, CROP_RIGHT);

                if (bitmapPreview.Width < MIN_IMAGE_SIZE || bitmapPreview.Height < MIN_IMAGE_SIZE)
                    bitmapPreview = RoundWithWhite(bitmapPreview);

                bitmapPreview.Save($@"C:\temp\{SheetName}.png");
                Graphics g = Graphics.FromImage(bitmapPreview);
            }

            return ToBitmapImage(bitmapPreview);
        }

        public BitmapImage SmallPreview
        {
            get
            {
                var previewImage = Preview();
                if (previewImage == null) return new BitmapImage();
                
                var preview = ToImage(previewImage) as Bitmap;
                return ToBitmapImage(preview.Clone(new System.Drawing.Rectangle(0, 0, MIN_IMAGE_SIZE, MIN_IMAGE_SIZE), preview.PixelFormat));
            }
        }

        public void ExportToPDF(string filePath)
        {
            excelSheetInterop.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, filePath, XlFixedFormatQuality.xlQualityStandard, true, true, 1, 10, false);
        }

        private static Bitmap CropImage(Image originalImage, int cropTop, int cropBottom, int cropLeft, int cropRight)
        {
            Bitmap cropImage = new Bitmap(originalImage);

            Bitmap bmpCrop = cropImage.Clone(new System.Drawing.Rectangle(cropLeft, cropTop, cropImage.Width - cropLeft - cropRight, cropImage.Height - cropTop - cropBottom), cropImage.PixelFormat);

            return bmpCrop;
        }

        public static BitmapImage ToBitmapImage(Bitmap bitmap)
        {
            using (var memory = new MemoryStream())
            {
                bitmap.Save(memory, ImageFormat.Png);
                memory.Position = 0;

                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();
                bitmapImage.Freeze();

                return bitmapImage;
            }
        }

        public static Bitmap RoundWithWhite(Bitmap bitmap)
        {
            Bitmap newImage = new Bitmap(bitmap.Width > MIN_IMAGE_SIZE ? bitmap.Width : MIN_IMAGE_SIZE, bitmap.Height > MIN_IMAGE_SIZE ? bitmap.Height : MIN_IMAGE_SIZE);
            using (Graphics graphics = Graphics.FromImage(newImage))
            {
                graphics.Clear(Color.White);
                graphics.DrawImage(bitmap, 0, 0);
            }

            return newImage;
        }

        public static Image ToImage(BitmapImage bitmap)
        {
            var image = BitmapImageToBitmap(bitmap) as Image;
            return image;
        }

        public static Bitmap BitmapImageToBitmap(BitmapImage bitmapImage)
        {
            using (MemoryStream outStream = new MemoryStream())
            {
                BitmapEncoder enc = new BmpBitmapEncoder();
                enc.Frames.Add(BitmapFrame.Create(bitmapImage));
                enc.Save(outStream);
                System.Drawing.Bitmap bitmap = new System.Drawing.Bitmap(outStream);

                return new Bitmap(bitmap);
            }
        }
    }
}
