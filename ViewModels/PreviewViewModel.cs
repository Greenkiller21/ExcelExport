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
    class PreviewViewModel : MvxViewModel
    {
        private ExcelFile _excelFile;
        private BitmapImage _currentPreview;
        private string _currentPreviewName;
        private List<string> _currentPreviewNameList = new List<string>();

        private bool isRendered = true;

        public ICommand Previous => new Command(() => 
        {
            isRendered = true;
            Render(GetCurrentSheetIndex() - 1);
        });

        public ICommand Next => new Command(() =>
        {
            isRendered = true;
            Render(GetCurrentSheetIndex() + 1);
        });

        public ICommand PreviewClicked => new Command((obj) => 
        {
            CurrentPreviewName = obj as string;
        });

        public ICommand Export => new Command(() =>
        {
            foreach (var sheet in ExcelFile.ExcelSheets)
            {
                if (sheet.ToExport)
                    sheet.ExportToPDF(SettingsViewModel.GetFileName(sheet));
            }
        });

        public ICommand SelectAll => new Command(() =>
        {
            foreach (var sheet in ExcelFile.ExcelSheets)
            {
                sheet.ToExport = true;
            }
        });

        public ICommand UnselectAll => new Command(() =>
        {
            foreach (var sheet in ExcelFile.ExcelSheets)
            {
                sheet.ToExport = false;
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

                if (!isRendered)
                {
                    isRendered = true;
                    Render(GetCurrentSheetIndex());
                }
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

        public List<ExcelSheet> ExcelSheets
        {
            get => ExcelFile?.ExcelSheets;
        }

        public bool Render(int index)
        {
            if (ExcelFile?.ExcelSheets == null || ExcelFile.ExcelSheets.Count < 1 || index < 0 || index >= ExcelFile.ExcelSheets.Count)
            {
                isRendered = false;
                return false;
            }

            var sheet = ExcelFile.ExcelSheets[index];
            CurrentPreview = sheet.BitmapPreview;
            CurrentPreviewName = sheet.SheetName;

            isRendered = false;
            return true;
        }

        public int GetCurrentSheetIndex()
        {
            return ExcelFile.ExcelSheets.IndexOf(ExcelFile.ExcelSheets.Where(file => file.SheetName == CurrentPreviewName).FirstOrDefault());
        }

        public override Task Initialize()
        {
            ExcelFile = Views.CustomComponents.Shared.CurrentExcelFile;

            foreach (var sheet in ExcelFile.ExcelSheets)
            {
                CurrentPreviewNameList.Add(sheet.SheetName);
            }

            Render(0);

            return base.Initialize();
        }
    }
}
