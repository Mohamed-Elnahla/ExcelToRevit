using Autodesk.Revit.Attributes;
using ExcelToRevit.ViewModels;
using ExcelToRevit.Views;
using Nice3point.Revit.Toolkit.External;

namespace ExcelToRevit.Commands
{
    [UsedImplicitly]
    [Transaction(TransactionMode.Manual)]
    public class Command : ExternalCommand
    {
        public override void Execute()
        {
            var viewModel = new ExcelToRevitViewModel();
            var view = new ExcelToRevitView(viewModel);
            view.ShowDialog();
        }
    }
}