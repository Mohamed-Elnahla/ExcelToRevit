using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using OfficeOpenXml;

namespace Elnahla.ExcelToRevit.Commands;

[Transaction(TransactionMode.Manual)]
internal class PilesPlacing : IExternalCommand
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
            //Family family = null;
            //String FileName = "";

            //convert Units Function
            const double _inchToMm = 25.4;
            const double _footToMm = 12 * _inchToMm;

            // Convert a given length in millimetres to feet.
            double MmToFoot(double length)
            {
                return length / _footToMm;
            }

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
                using (var PlacingPiles = new Transaction(doc, "Placing Piles"))
                {
                    //Working on all Data in first 5 columns in first sheet
                    for (var row = 2; row < 9999; row++)
                    {
                        PlacingPiles.Start();
                        if ((string)workSheet.Cells[row, 1].Value == null ||
                            (string)workSheet.Cells[row, 1].Value == "") break;

                        // Selecting Family Symbol
                        var Family_Symbol = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_StructuralFoundation).OfClass(typeof(FamilySymbol))
                            .OfType<FamilySymbol>().Where(x =>
                                x.FamilyName == "Pile-Concrete" && x.Name == (string)workSheet.Cells[row, 6].Value)
                            .FirstOrDefault();

                        //Making Family Symbol Active
                        if (Family_Symbol.IsActive == false)
                        {
                            Family_Symbol.Activate();
                            doc.Regenerate();
                        }


                        //Placing Piles
                        var location = new XYZ(MmToFoot((double)workSheet.Cells[row, 2].Value),
                            MmToFoot((double)workSheet.Cells[row, 3].Value),
                            MmToFoot((double)workSheet.Cells[row, 4].Value));
                        var Pile = doc.Create.NewFamilyInstance(location, Family_Symbol, StructuralType.Footing);
                        var InstanceID = Pile.GetTypeId();
                        var depth = Pile.LookupParameter("Depth");
                        depth.SetValueString(workSheet.Cells[row, 5].Value.ToString());
                        depth.Set(InstanceID);
                        var PileID = Pile.LookupParameter("Pile ID");
                        PileID.Set(workSheet.Cells[row, 1].Value.ToString());

                        PlacingPiles.Commit();
                    }
                }

                TaskDialog.Show("Revit", "New Piles Placed Successfully");
                return Result.Succeeded;
            }
        }
        catch (Exception e)
        {
            message = e.Message;
            return Result.Failed;
        }
    }
}