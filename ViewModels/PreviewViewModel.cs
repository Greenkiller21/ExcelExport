using ExcelExport.Excel;
using ExcelExport.Utils;
using Microsoft.Win32;
using MvvmCross.ViewModels;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Media.Imaging;

namespace ExcelExport.ViewModels
{
    class PreviewViewModel : MvxViewModel<ExcelFile>
    {
        private ExcelFile _excelFile;
        private List<ExcelSheet> _sheetsToExport = new List<ExcelSheet>();
        private BitmapImage _currentPreview;
        private string _currentPreviewName;
        private List<string> _currentPreviewNameList = new List<string>();

        private int currentPreviewIndex = 0;

        public ICommand Previous => new Command(() => 
        {
            currentPreviewIndex--;
            if (!Render(currentPreviewIndex))
            {
                currentPreviewIndex++;
            }
        });

        public ICommand Next => new Command(() =>
        {
            currentPreviewIndex++;
            if (!Render(currentPreviewIndex))
            {
                currentPreviewIndex--;
            }
        });

        public ExcelFile ExcelFile
        {
            get => _excelFile;
            set => SetProperty(ref _excelFile, value);
        }

        public string CurrentPreviewName
        {
            get => _currentPreviewName;
            set
            {
                SetProperty(ref _currentPreviewName, value);
                Render(ExcelFile.ExcelSheets.IndexOf(ExcelFile.ExcelSheets.Where(file => file.SheetName == value).FirstOrDefault()));
            }
        }

        public List<string> CurrentPreviewNameList
        {
            get => _currentPreviewNameList;
            set => SetProperty(ref _currentPreviewNameList, value);
        }

        public BitmapImage CurrentPreview
        {
            get => _currentPreview;
            set => SetProperty(ref _currentPreview, value);
        }

        public bool Render(int index)
        {
            if (ExcelFile?.ExcelSheets == null || ExcelFile.ExcelSheets.Count < 1 || index < 0 || index >= ExcelFile.ExcelSheets.Count) 
                return false;


            var sheet = ExcelFile.ExcelSheets[index];
            CurrentPreview = sheet.Preview();

            return true;
        }

        public override void Prepare(ExcelFile parameter)
        {
            ExcelFile = parameter;
            
            foreach(var sheet in ExcelFile.ExcelSheets)
            {
                CurrentPreviewNameList.Add(sheet.SheetName);
            }

            CurrentPreviewName = ExcelFile.ExcelSheets[0].SheetName;
            Render(0);
        }
    }
}
