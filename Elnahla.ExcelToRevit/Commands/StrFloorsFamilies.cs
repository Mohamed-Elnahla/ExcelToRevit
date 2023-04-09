using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;

namespace Elnahla.ExcelToRevit.Commands;

[Transaction(TransactionMode.Manual)]
internal class StrFloorsFamilies : IExternalCommand
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
            var FloorType = collector.OfClass(typeof(FloorType))
                .WhereElementIsElementType()
                .Cast<FloorType>()
                .First(x => x.Name == "Generic 300mm");
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
                        var newFloorType = FloorType.Duplicate((string)workSheet.Cells[row, 1].Value) as FloorType;
                        var newFloorTypeStr = newFloorType.GetCompoundStructure();
                        var layers = newFloorTypeStr.GetLayers();
                        var layerwidth = MmToFoot((double)workSheet.Cells[row, 2].Value);
                        var layerindex = newFloorTypeStr.GetFirstCoreLayerIndex();
                        var cs_layers = newFloorTypeStr.GetLayers();
                        foreach (var cs_layer in cs_layers)
                        {
                            newFloorTypeStr.SetLayerWidth(layerindex, layerwidth);
                            layerindex++;
                        }

                        newFloorType.SetCompoundStructure(newFloorTypeStr);
                        newFamilyTypeTransaction.Commit();
                    }
                }
            }

            TaskDialog.Show("Revit", "New Structural Floors Types Created Successfully");
            return Result.Succeeded;
        }
        catch (Exception e)
        {
            message = e.Message;
            return Result.Failed;
        }
    }
}