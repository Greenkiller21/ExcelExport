using ExcelExport.ViewModels;
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
            //Mvx.IoCProvider.RegisterType<ICalculationService, CalculationService>();
            RegisterAppStart<MainViewModel>();
        }
    }
}
