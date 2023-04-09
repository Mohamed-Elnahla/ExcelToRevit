using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using OfficeOpenXml;

namespace Elnahla.ExcelToRevit.Commands;

[Transaction(TransactionMode.Manual)]
internal class ColumnsPlacing : IExternalCommand
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
                using (var Placing_Columns = new Transaction(doc, "Placing Columns"))
                {
                    //Working on all Data in first 5 columns in first sheet
                    for (var row = 2; row < 9999; row++)
                    {
                        Placing_Columns.Start();
                        if ((string)workSheet.Cells[row, 1].Value == null ||
                            (string)workSheet.Cells[row, 1].Value == "") break;

                        // Selecting Family Symbol
                        var Family_Symbol = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_StructuralColumns).OfClass(typeof(FamilySymbol))
                            .OfType<FamilySymbol>().Where(x =>
                                x.FamilyName == "Conc.Columns" && x.Name == (string)workSheet.Cells[row, 5].Value)
                            .FirstOrDefault();

                        //Making Family Symbol Active
                        if (Family_Symbol.IsActive == false)
                        {
                            Family_Symbol.Activate();
                            doc.Regenerate();
                        }


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

                        //Setting Location
                        var location = new XYZ(MmToFoot((double)workSheet.Cells[row, 2].Value),
                            MmToFoot((double)workSheet.Cells[row, 3].Value),
                            MmToFoot((double)workSheet.Cells[row, 4].Value));

                        //Placing Columns
                        var Column = doc.Create.NewFamilyInstance(location, Family_Symbol, BaseLevel,
                            StructuralType.Column);

                        //Select Top Level by Elevation and create one if not existing
                        var TopLevelEle = MmToFoot((double)workSheet.Cells[row, 7].Value);
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

                        var TopLevelR = Column.LookupParameter("Top Level");
                        TopLevelR.Set(TopLevelId);

                        //Setting offset to Zero
                        var B_off = Column.LookupParameter("Base Offset");
                        B_off.Set(0);
                        var T_off = Column.LookupParameter("Top Offset");
                        T_off.Set(0);


                        //Setting Orientation
                        var startPoint = new XYZ(MmToFoot((double)workSheet.Cells[row, 2].Value),
                            MmToFoot((double)workSheet.Cells[row, 3].Value),
                            MmToFoot((double)workSheet.Cells[row, 4].Value));
                        var newZ = MmToFoot((double)workSheet.Cells[row, 3].Value) + 1000;
                        var endPoint = new XYZ(MmToFoot((double)workSheet.Cells[row, 2].Value),
                            MmToFoot((double)workSheet.Cells[row, 3].Value), newZ);
                        var Axis = Line.CreateBound(startPoint, endPoint);
                        var ID = Column.Id;
                        var Angle = (double)workSheet.Cells[row, 6].Value * Math.PI / 180;
                        ElementTransformUtils.RotateElement(doc, ID, Axis, Angle);
                        Placing_Columns.Commit();
                    }
                }

                TaskDialog.Show("Revit", "New Columns Placed Successfully");
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