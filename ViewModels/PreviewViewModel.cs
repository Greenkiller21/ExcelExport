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

namespace ExcelExport.ViewModels
{
    class PreviewViewModel : MvxViewModel<ExcelFile>
    {
        private ExcelFile _excelFile;
        private List<ExcelSheet> _sheetsToExport = new List<ExcelSheet>();
        public Bitmap _currentPreview;

        public ExcelFile ExcelFile
        {
            get => _excelFile;
            set => SetProperty(ref _excelFile, value);
        }

        private string _test;
        public string test
        {
            get => _test;
            set => SetProperty(ref _test, value);
        }

        public Bitmap CurrentPreview
        {
            get => _currentPreview;
            set => SetProperty(ref _currentPreview, value);
        }

        public override void Prepare(ExcelFile parameter)
        {
            ExcelFile = parameter;
            test = parameter.ExcelSheets[0].SheetName;
            CurrentPreview = ExcelFile.ExcelSheets[0].Preview();
        }
    }
}
