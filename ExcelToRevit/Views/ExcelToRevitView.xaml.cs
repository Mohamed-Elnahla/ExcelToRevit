using ExcelToRevit.ViewModels;

namespace ExcelToRevit.Views
{
    public partial class ExcelToRevitView
    {
        public ExcelToRevitView(ExcelToRevitViewModel viewModel)
        {
            InitializeComponent();
            DataContext = viewModel;
        }
    }
}