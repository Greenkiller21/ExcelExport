using ExcelExport.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Spire.Xls;
using System.Drawing;
using System.IO;
using MvvmCross.Platforms.Wpf.Views;

namespace ExcelExport
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MvxWindow
    {
        Excel.ExcelFile test;

        public MainWindow()
        {
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(OnProcessExit);

            DataContext = new MainViewModel();
            InitializeComponent();

            ExcelTest();
        }

        private void OnProcessExit(object sender, EventArgs e)
        {
            test.Close();
        }

        private void ExcelTest()
        {
            test = new Excel.ExcelFile(@"C:\temp\Temp.xlsx");

            System.Drawing.Image imageTest = test.ExcelSheets[0].Preview();

            imageTest.Save(@"C:\temp\ImageTest.jpeg");
        }
    }
}
