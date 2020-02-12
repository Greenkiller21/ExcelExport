using ExcelExport.ViewModels;
using MvvmCross;
using MvvmCross.Core.Navigation;
using MvvmCross.ViewModels;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelExport
{
    public class AppStart : MvxApplication
    {
        public override void Initialize()
        {
            RegisterAppStart<MainViewModel>();
        }
    }
}
