using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Elnahla.ExcelToRevit.Commands;

[Transaction(TransactionMode.Manual)]
internal class TitleblocksFamilies : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            //Get UIDocument
            var uidoc = commandData.Application.ActiveUIDocument;

            //Get Document
            var doc = uidoc.Document;

            //Declaring Variables
            Family family = null;
            Family family1 = null;
            Family family2 = null;
            Family family3 = null;
            Family family4 = null;

            //Family Path
            var FileName = "C:/Program Files/Elnahla Revit Plugin/Families/Titleblocks/A0 metric.rfa";
            var FileName1 = "C:/Program Files/Elnahla Revit Plugin/Families/Titleblocks/A1 metric.rfa";
            var FileName2 = "C:/Program Files/Elnahla Revit Plugin/Families/Titleblocks/A2 metric.rfa";
            var FileName3 = "C:/Program Files/Elnahla Revit Plugin/Families/Titleblocks/A3 metric.rfa";
            var FileName4 = "C:/Program Files/Elnahla Revit Plugin/Families/Titleblocks/A4 metric.rfa";

            using (var trans = new Transaction(doc, "Import Titleblocks Families"))
            {
                trans.Start();
                doc.LoadFamily(FileName, out family);
                doc.LoadFamily(FileName1, out family1);
                doc.LoadFamily(FileName2, out family2);
                doc.LoadFamily(FileName3, out family3);
                doc.LoadFamily(FileName4, out family4);
                trans.Commit();
            }

            TaskDialog.Show("Revit", "TileBlocks Families Loaded Successfully");
            return Result.Succeeded;
        }
        catch (Exception e)
        {
            message = e.Message;
            return Result.Failed;
        }
    }
}