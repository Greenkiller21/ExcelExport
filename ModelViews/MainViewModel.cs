using ExcelExport.Utils;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace ExcelExport.ModelViews
{
    class MainViewModel : BaseViewModel
    {
        private string _testText = "test";
        public string TestText
        {
            get => _testText;
            set => SetProperty(ref _testText, value);
        }

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

                string fileName = ofdFile.FileNames?.First();
                if (string.IsNullOrWhiteSpace(fileName)) return;


            });
        }
    }
}
