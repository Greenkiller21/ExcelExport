using ExcelExport.Excel;
using ExcelExport.Utils;
using Microsoft.Win32;
using MvvmCross.ViewModels;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
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
            bool exported = false;
            DialogResult drReplace = DialogResult.No;
            DialogResult drSaveChoice = DialogResult.No;

            foreach (var sheet in ExcelFile.ExcelSheets)
            {
                if (sheet.ToExport)
                {
                    string filePath = SettingsViewModel.GetFileName(sheet);

                    if (File.Exists(filePath))
                    {
                        if (drSaveChoice == DialogResult.No)
                        {
                            drReplace = MessageBox.Show(string.Format("The file named {0} already exists, do you want to replace it ?", filePath.Split('\\').Last()), "File already exists", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                            drSaveChoice = MessageBox.Show("Do you want to apply this choice for the next files ?", "File already exists", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                        }

                        if (drReplace == DialogResult.Yes)
                        {
                            sheet.ExportToPDF(filePath);
                        }
                    }
                    else
                    {
                        sheet.ExportToPDF(filePath);
                    }
                    
                    exported = true;
                }
            }

            if(exported)
            {
                MessageBox.Show("The export has been successfully completed", "Exportation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else 
            {
                MessageBox.Show("There is no file selected for the export", "Exportation", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        });

        public ICommand SelectAll => new Command(() =>
        {
            foreach (var sheet in ExcelFile.ExcelSheets)
            {
                sheet.ToExport = true;
            }

            RaisePropertyChanged(nameof(ExcelFile.ExcelSheets));
        });

        public ICommand UnselectAll => new Command(() =>
        {
            foreach (var sheet in ExcelFile.ExcelSheets)
            {
                sheet.ToExport = false;
            }

            RaisePropertyChanged(nameof(ExcelFile.ExcelSheets));
        });

        public ICommand DestinationFolder => new Command(() =>
        {
            if (Directory.Exists(SettingsViewModel.GetExportFolder()))
            {
                Process.Start("explorer.exe", SettingsViewModel.GetExportFolder());
            }
            else
            {
                MessageBox.Show("The folder doesn't exist or isn't configured", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });

        public ExcelFile ExcelFile
        {
            get => _excelFile;
            set 
            {
                SetProperty(ref _excelFile, value);
                ExcelFile.RefreshAction = () => { RaisePropertyChanged(nameof(ExcelSheets)); };
            }
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

        public MvxObservableCollection<ExcelSheet> ExcelSheets
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
