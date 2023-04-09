using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;

namespace Elnahla.ExcelToRevit.Commands;

[Transaction(TransactionMode.Manual)]
internal class WallsPlacing : IExternalCommand
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
                using (var Placing_Walls = new Transaction(doc, "Placing Walls"))
                {
                    //Working on all Data in first 5 columns in first sheet
                    for (var row = 2; row < 9999; row++)
                    {
                        Placing_Walls.Start();
                        if ((string)workSheet.Cells[row, 1].Value == null ||
                            (string)workSheet.Cells[row, 1].Value == "") break;

                        // Selecting Family Symbol
                        var wall_Type = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls)
                            .OfClass(typeof(WallType)).OfType<WallType>().Where(x =>
                                x.FamilyName == "Basic Wall" && x.Name == (string)workSheet.Cells[row, 5].Value)
                            .FirstOrDefault();
                        var wallTypeId = wall_Type.Id;

                        ////Making Family Symbol Active
                        //if (Family_Symbol.IsActive == false)
                        //{
                        //    Family_Symbol.Activate();
                        //    doc.Regenerate();
                        //}

                        //Select Base Level by Elevation and create one if not existing
                        var BaseLevelEle = MmToFoot((double)workSheet.Cells[row, 4].Value);
                        var BaseLevel = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels)
                            .OfClass(typeof(Level)).OfType<Level>().Where(x =>
                                x.Elevation - BaseLevelEle < 0.01 && x.Elevation - BaseLevelEle > -0.01)
                            .FirstOrDefault();
                        if (BaseLevel == null)
                        {
                            // Begin in creating a level
                            BaseLevel = Level.Create(doc, BaseLevelEle);
                            if (null == BaseLevel) throw new Exception("Create a new level failed.");
                        }

                        var BaseLevelId = BaseLevel.Id;

                        //Path of wall
                        var pt1 = new XYZ(MmToFoot((double)workSheet.Cells[row, 2].Value),
                            MmToFoot((double)workSheet.Cells[row, 3].Value),
                            MmToFoot((double)workSheet.Cells[row, 4].Value));
                        var pt2 = new XYZ(MmToFoot((double)workSheet.Cells[row, 6].Value),
                            MmToFoot((double)workSheet.Cells[row, 7].Value),
                            MmToFoot((double)workSheet.Cells[row, 4].Value));
                        var line = Line.CreateBound(pt1, pt2);

                        //Placing Walls
                        var Walls = Wall.Create(doc, line, BaseLevelId, true);


                        //Select Top Level by Elevation and create one if not existing
                        var TopLevelEle = MmToFoot((double)workSheet.Cells[row, 8].Value);
                        var TopLevel = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels)
                            .OfClass(typeof(Level)).OfType<Level>().Where(x =>
                                x.Elevation - TopLevelEle < 0.01 && x.Elevation - TopLevelEle > -0.01).FirstOrDefault();
                        if (TopLevel == null)
                        {
                            // Begin in creating a level
                            TopLevel = Level.Create(doc, TopLevelEle);
                            if (null == TopLevel) throw new Exception("Create a new level failed.");
                        }

                        var TopLevelId = TopLevel.Id;

                        var TopLevelR = Walls.LookupParameter("Top Constraint");
                        TopLevelR.Set(TopLevelId);

                        //Setting offset to Zero
                        var B_off = Walls.LookupParameter("Base Offset");
                        B_off.Set(0);
                        var T_off = Walls.LookupParameter("Top Offset");
                        T_off.Set(0);

                        //Set wall type
                        Walls.ChangeTypeId(wallTypeId);

                        Placing_Walls.Commit();
                    }
                }

                TaskDialog.Show("Revit", "New Walls Placed Successfully");
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