using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;

namespace Elnahla.ExcelToRevit.Commands;

[Transaction(TransactionMode.Manual)]
public class CreatingLevels : IExternalCommand
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
            Level level;

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

                //Working on all Data in first 2 columns in first sheet
                for (var row = 2; row < 9999; row++)
                {
                    //Get Level Name
                    var Levelname = (string)workSheet.Cells[row, 1].Value;
                    //NullReferenceException
                    if (Levelname == null || Levelname == "") break;
                    //Get Level Elevation
                    var LevelEle = MmToFoot((double)workSheet.Cells[row, 2].Value);

                    //Creating Level
                    using (var trans = new Transaction(doc, "Creating Level" + Levelname))
                    {
                        trans.Start();
                        // Begin to create a level
                        level = Level.Create(doc, LevelEle);
                        if (null == level) throw new Exception("Create a new level failed.");

                        // Change the level name
                        level.Name = Levelname;
                        trans.Commit();
                    }

                    //Creating Structrual View Plan
                    using (var trans = new Transaction(doc, "Creating View Plan"))
                    {
                        trans.Start();
                        var levelId = level.Id;
                        var vft = new FilteredElementCollector(doc).OfClass(typeof(ViewFamilyType))
                            .Cast<ViewFamilyType>().FirstOrDefault(x => ViewFamily.StructuralPlan == x.ViewFamily);
                        var floorPlanId = vft.Id;
                        var floorView = ViewPlan.Create(doc, floorPlanId, levelId);
                        trans.Commit();
                    }
                }
            }

            return Result.Succeeded;
        }
        catch (Exception e)
        {
            message = e.Message;
            return Result.Failed;
        }
    }
}