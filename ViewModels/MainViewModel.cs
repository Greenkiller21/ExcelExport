using ExcelExport.Excel;
using ExcelExport.Utils;
using Microsoft.Win32;
using MvvmCross.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ExcelExport.ViewModels
{
    class MainViewModel : MvxViewModel
    {
        private ExcelFile file;

        public ICommand OpenFiles
        {
            get => new Command(() => 
            {
                OpenFileDialog ofdFile = new OpenFileDialog();
                ofdFile.Filter = "Excel files (*.xlsx)|*.xlsx";
                ofdFile.Title = "Choose a file to open";
                ofdFile.CheckFileExists = true;
                ofdFile.CheckPathExists = true;
                ofdFile.Multiselect = false;
                ofdFile.InitialDirectory = Environment.ExpandEnvironmentVariables(@"%USERPROFILE%\Downloads");
                if (ofdFile.ShowDialog() != true) return;

                string filePath = ofdFile.FileNames?.First();
                if (string.IsNullOrWhiteSpace(filePath)) return;

                file = new ExcelFile(filePath);
            });
        }
    }
}
