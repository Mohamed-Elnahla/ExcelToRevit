using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using OfficeOpenXml;

namespace Elnahla.ExcelToRevit.Commands;

[Transaction(TransactionMode.Manual)]
internal class BeamsPlacing : IExternalCommand
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
                using (var Placing_Beams = new Transaction(doc, "Placing Beams"))
                {
                    //Working on all Data in first 5 columns in first sheet
                    for (var row = 2; row < 9999; row++)
                    {
                        Placing_Beams.Start();
                        if ((string)workSheet.Cells[row, 1].Value == null ||
                            (string)workSheet.Cells[row, 1].Value == "") break;

                        // Selecting Family Symbol
                        var Family_Symbol = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol))
                            .OfType<FamilySymbol>().Where(x =>
                                x.FamilyName == "Conc.Beams" && x.Name == (string)workSheet.Cells[row, 5].Value)
                            .FirstOrDefault();

                        //Making Family Symbol Active
                        if (Family_Symbol.IsActive == false)
                        {
                            Family_Symbol.Activate();
                            doc.Regenerate();
                        }

                        //Select Level by Elevation and create one if not existing
                        var LevelEle = MmToFoot((double)workSheet.Cells[row, 4].Value);
                        var Level = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels)
                            .OfClass(typeof(Level)).OfType<Level>()
                            .Where(x => x.Elevation - LevelEle < 0.01 && x.Elevation - LevelEle > -0.01)
                            .FirstOrDefault();
                        if (Level == null)
                        {
                            // Begin in creating a level
                            Level = Level.Create(doc, LevelEle);
                            if (null == Level) throw new Exception("Create a new level failed.");
                        }

                        //Path of Beam
                        var startPoint = new XYZ(MmToFoot((double)workSheet.Cells[row, 2].Value),
                            MmToFoot((double)workSheet.Cells[row, 3].Value),
                            MmToFoot((double)workSheet.Cells[row, 4].Value));
                        var endPoint = new XYZ(MmToFoot((double)workSheet.Cells[row, 6].Value),
                            MmToFoot((double)workSheet.Cells[row, 7].Value),
                            MmToFoot((double)workSheet.Cells[row, 4].Value));
                        var Axis = Line.CreateBound(startPoint, endPoint);

                        //Placing Walls
                        var Beam = doc.Create.NewFamilyInstance(Axis, Family_Symbol, Level, StructuralType.Beam);

                        //Setting offset to Zero
                        var S_off = Beam.LookupParameter("Start Level Offset");
                        S_off.Set(0);
                        var E_off = Beam.LookupParameter("End Level Offset");
                        E_off.Set(0);

                        Placing_Beams.Commit();
                    }
                }

                TaskDialog.Show("Revit", "New Beams Placed Successfully");
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