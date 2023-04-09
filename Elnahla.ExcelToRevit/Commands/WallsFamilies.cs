using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;

namespace Elnahla.ExcelToRevit.Commands;

[Transaction(TransactionMode.Manual)]
internal class WallsFamilies : IExternalCommand
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
            var Excelfile = "";

            //convert Units Function
            const double _inchToMm = 25.4;
            const double _footToMm = 12 * _inchToMm;

            // Convert a given length in millimetres to feet.
            double MmToFoot(double length)
            {
                return length / _footToMm;
            }

            //Get Family Symbol
            var collector = new FilteredElementCollector(doc);
            var wallType = collector.OfClass(typeof(WallType))
                .WhereElementIsElementType()
                .Cast<WallType>()
                .First(x => x.Name == "Foundation - 300mm Concrete");
            //Open file Dialog Box
            var dialogRead = new OpenFileDialog();
            dialogRead.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            dialogRead.ShowDialog();

            //Store Excel File into Variable
            Excelfile = dialogRead.FileName;

            //Using Data From Excel File
            using (var ExcelPack = new ExcelPackage(new FileInfo(Excelfile)))
            {
                //Get WorkSheet
                var workSheet = ExcelPack.Workbook.Worksheets[1];
                // Start transaction for the family document
                using (var newFamilyTypeTransaction = new Transaction(doc, "Add Type to Family"))
                {
                    //Working on all Data in first 2 columns in first sheet
                    for (var row = 2; row < 9999; row++)
                    {
                        newFamilyTypeTransaction.Start();
                        if ((string)workSheet.Cells[row, 1].Value == null ||
                            (string)workSheet.Cells[row, 1].Value == "") break;
                        var newWallType = wallType.Duplicate((string)workSheet.Cells[row, 1].Value) as WallType;
                        var newWallTypeStr = newWallType.GetCompoundStructure();
                        var layers = newWallTypeStr.GetLayers();
                        var layerwidth = MmToFoot((double)workSheet.Cells[row, 2].Value);
                        var layerindex = newWallTypeStr.GetFirstCoreLayerIndex();
                        var cs_layers = newWallTypeStr.GetLayers();
                        foreach (var cs_layer in cs_layers)
                        {
                            newWallTypeStr.SetLayerWidth(layerindex, layerwidth);
                            layerindex++;
                        }

                        newWallType.SetCompoundStructure(newWallTypeStr);
                        newFamilyTypeTransaction.Commit();
                    }
                }
            }

            TaskDialog.Show("Revit", "New Walls Types Created Successfully");
            return Result.Succeeded;
        }
        catch (Exception e)
        {
            message = e.Message;
            return Result.Failed;
        }
    }
}