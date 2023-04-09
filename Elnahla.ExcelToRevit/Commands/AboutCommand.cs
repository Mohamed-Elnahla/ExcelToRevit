using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Elnahla.ExcelToRevit.Views;

namespace Elnahla.ExcelToRevit.Commands;

[Transaction(TransactionMode.Manual)]
public class AboutCommand : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        var aboutView = new AboutView();
        aboutView.ShowDialog();

        return Result.Succeeded;
    }
}