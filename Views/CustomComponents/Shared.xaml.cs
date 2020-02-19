using ExcelExport.Excel;
using ExcelExport.Utils;
using ExcelExport.ViewModels;
using Microsoft.Win32;
using MvvmCross;
using MvvmCross.Navigation;
using MvvmCross.ViewModels;
using System;
using System.Collections.Generic;
using System.IO;
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
        public static ExcelFile CurrentExcelFile = null;

        public static ICommand OpenFile => new Command(async () =>
        {
            var config = Mvx.IoCProvider.GetSingleton<ConfigFile>();
            var efc = config["Settings"]["ExcelFolder"].Value;

            var ofdFile = new OpenFileDialog();
            ofdFile.Filter = "Excel Files|*.xlsx;*.xlsm;*.xlsb;*.xltx;*.xltm;*.xls;*.xlt;*.xls;*.xml;*.xml;*.xlam;*.xla;*.xlw;*.xlr";
            ofdFile.Title = "Choose a file to open";
            ofdFile.CheckFileExists = true;
            ofdFile.CheckPathExists = true;
            ofdFile.Multiselect = false;
            ofdFile.InitialDirectory = string.IsNullOrWhiteSpace(efc) || !Directory.Exists(efc) ? Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Downloads") : efc;
            if (ofdFile.ShowDialog() != true) 
                return;

            var filePath = ofdFile.FileNames?.First();
            if (string.IsNullOrWhiteSpace(filePath)) 
                return;

            config["Settings"]["ExcelFolder"].Value = new FileInfo(filePath).DirectoryName;
            config.Save();

            FileChanged(null, null);

            CurrentExcelFile = new ExcelFile(filePath);

            await Mvx.IoCProvider.Resolve<IMvxNavigationService>().Navigate<PreviewViewModel>();
        });

        public static void FileChanged(object sender, EventArgs e)
        {
            AppDomain.CurrentDomain.ProcessExit -= FileChanged;
            CurrentExcelFile?.Close();
        }

        public static ICommand Settings => new Command(async () =>
        {
            await Mvx.IoCProvider.Resolve<IMvxNavigationService>().Navigate<SettingsViewModel>();
        });

        public Shared()
        {
            AppDomain.CurrentDomain.ProcessExit += FileChanged;
            InitializeComponent();
        }
    }
}
