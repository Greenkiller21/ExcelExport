﻿using ExcelExport.Excel;
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
    class MainViewModel : MvxViewModel
    {
        public ICommand OpenFile => Views.CustomComponents.Shared.OpenFile;
    }
}
