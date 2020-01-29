using ExcelExport.Utils;
using ExcelExport.ViewModels;
using MvvmCross;
using MvvmCross.Core;
using MvvmCross.Navigation;
using MvvmCross.Platforms.Wpf.Core;
using MvvmCross.Platforms.Wpf.Views;
using MvvmCross.ViewModels;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelExport
{
    /// <summary>
    /// Logique d'interaction pour App.xaml
    /// </summary>
    public partial class App : MvvmCross.Platforms.Wpf.Views.MvxApplication
    {
        protected override void RegisterSetup()
        {
            this.RegisterSetupType<MvxWpfSetup<AppStart>>();
        }

        public override void ApplicationInitialized()
        {
            base.ApplicationInitialized();

            var config = new ConfigFile(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelExport", "config.ini"));
            config.Load();
            Mvx.IoCProvider.RegisterSingleton(config);

            SettingsViewModel.Load();

            Stack<Type> navigation = new Stack<Type>(new Type[] { typeof(MainViewModel) });
            Mvx.IoCProvider.Resolve<IMvxNavigationService>().BeforeNavigate += (sender, e) =>
            {
                if (navigation.Count > 0 && e.ViewModel.GetType() == navigation.Peek() && e.ViewModel.GetType() != typeof(PreviewViewModel))
                {
                    e.Cancel = true;
                }
                else
                {
                    navigation.Push(e.ViewModel.GetType());
                }
            };

            Mvx.IoCProvider.Resolve<IMvxNavigationService>().AfterClose += (sender, e) =>
            {
                navigation.Pop();
            };
        }
    }
}
