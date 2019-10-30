using ExcelExport.Excel;
using ExcelExport.Utils;
using Microsoft.Win32;
using MvvmCross;
using MvvmCross.Navigation;
using MvvmCross.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ExcelExport.ViewModels
{
    public class SettingsViewModel : MvxViewModel
    {
        private string _fileNaming = "{year}.{month}.{day}_{sheetName}.pdf";

        public string FileNaming
        {
            get => _fileNaming;
            set => SetProperty(ref _fileNaming, value);
        }

        public ICommand Back => new Command(() => 
        {
            Mvx.IoCProvider.Resolve<IMvxNavigationService>().Close(this);
        });
    }
}
