using ExcelExport.Excel;
using ExcelExport.Utils;
using ExcelExport.ViewModels;
using Microsoft.Win32;
using MvvmCross;
using MvvmCross.Navigation;
using MvvmCross.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;

namespace ExcelExport.Views.CustomComponents
{
    /// <summary>
    /// Logique d'interaction pour Shared.xaml
    /// </summary>
    public partial class Shared : UserControl
    {
        public static ICommand OpenFile => new Command(async () =>
        {
            var ofdFile = new OpenFileDialog();
            ofdFile.Filter = "Excel files (*.xlsx)|*.xlsx";
            ofdFile.Title = "Choose a file to open";
            ofdFile.CheckFileExists = true;
            ofdFile.CheckPathExists = true;
            ofdFile.Multiselect = false;
            ofdFile.InitialDirectory = Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Downloads");
            if (ofdFile.ShowDialog() != true) return;

            var filePath = ofdFile.FileNames?.First();
            if (string.IsNullOrWhiteSpace(filePath)) return;

            FileChangedEvent?.Invoke(null, null);

            var file = new ExcelFile(filePath);

            await Mvx.IoCProvider.Resolve<IMvxNavigationService>().Navigate<PreviewViewModel, ExcelFile>(file);
        });

        public static ICommand Settings => new Command(async () =>
        {
            await Mvx.IoCProvider.Resolve<IMvxNavigationService>().Navigate<SettingsViewModel>();
        });

        public Shared()
        {
            InitializeComponent();
        }

        public delegate void FileChangedEventHandler(object sender, EventArgs e);
        public static event FileChangedEventHandler FileChangedEvent;
    }
}
