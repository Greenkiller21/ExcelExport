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
using ExcelExport.Excel;
using MvvmCross.Navigation;
using MvvmCross;

namespace ExcelExport
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MvxWindow
    {
        public MainWindow()
        {
            InitializeComponent();
        }
    }
}
