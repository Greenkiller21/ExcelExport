using ExcelExport.Excel;
using ExcelExport.Utils;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using MvvmCross;
using MvvmCross.Navigation;
using MvvmCross.ViewModels;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using static System.Environment;

namespace ExcelExport.ViewModels
{
    public class SettingsViewModel : MvxViewModel
    {
        private static string _fileNaming = "{year}.{month}.{day}_{sheetName}.pdf";
        private static string _exportFolder = GetFolderPath(SpecialFolder.MyDocuments);

        public SettingsViewModel()
        {
            var config = Mvx.IoCProvider.GetSingleton<ConfigFile>();
            var efc = config["Settings"]["ExportFolder"].Value;

            ExportFolder = efc;
        }

        public string FileNaming
        {
            get => _fileNaming;
            set => SetProperty(ref _fileNaming, value);
        }

        public string ExportFolder
        {
            get => _exportFolder;
            set
            {
                var config = Mvx.IoCProvider.GetSingleton<ConfigFile>();

                if (!string.IsNullOrWhiteSpace(value) && Directory.Exists(value))
                {
                    SetProperty(ref _exportFolder, value);
                    config["Settings"]["ExportFolder"].Value = value;
                    config.Save();
                }
            }
        }

        public static string GetExportFolder()
        {
            return _exportFolder;
        }

        public static string GetFileName(ExcelSheet sheet)
        {
            var dicParams = new Dictionary<string, string>()
            {
                { "year", DateTime.Today.Year.ToString("0000") },
                { "month", DateTime.Today.Month.ToString("00") },
                { "day", DateTime.Today.Day.ToString("00") },
                { "sheetName", sheet.SheetName }
            };

            var fileName = _fileNaming;

            foreach (var truc in dicParams)
            {
                fileName = fileName.Replace("{" + truc.Key + "}", truc.Value);
            }

            return Path.Combine(GetExportFolder(), fileName);
        }

        public ICommand Back => new Command(() => 
        {
            Mvx.IoCProvider.Resolve<IMvxNavigationService>().Close(this);
        });

        public ICommand ChooseLocation => new Command(() =>
        {
            var dialog = new CommonOpenFileDialog("Choose a folder ...");
            dialog.IsFolderPicker = true;
            var result = dialog.ShowDialog();
            dialog.InitialDirectory = ExportFolder;

            if (result == CommonFileDialogResult.Ok)
            {
                ExportFolder = dialog.FileName;
            }
        });
    }
}
