using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
