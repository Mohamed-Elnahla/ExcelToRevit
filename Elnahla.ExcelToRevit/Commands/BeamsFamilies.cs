using System.IO;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;

namespace Elnahla.ExcelToRevit.Commands;

[TransactionAttribute(TransactionMode.Manual)]
internal class BeamsFamilies : IExternalCommand
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
            Family family = null;
            var FileName = "";

            //convert Units Function
            const double _inchToMm = 25.4;
            const double _footToMm = 12 * _inchToMm;

            // Convert a given length in millimetres to feet.
            double MmToFoot(double length)
            {
                return length / _footToMm;
            }


            //Family Path
            FileName = "C:/Program Files/Elnahla Revit Plugin/Families/Conc.Beams.rfa";

            using (var trans = new Transaction(doc, "Import Beams Family"))
            {
                trans.Start();
                doc.LoadFamily(FileName, out family);
                trans.Commit();
            }

            // Get Family document for family
            var familyDoc = doc.EditFamily(family);
            if (null == familyDoc) TaskDialog.Show("Revit", "could not open a family for edit");

            var familyManager = familyDoc.FamilyManager;
            if (null == familyManager) TaskDialog.Show("Revit", "could not get a family manager");
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
                using (var newFamilyTypeTransaction = new Transaction(familyDoc, "Add Type to Family"))
                {
                    //Working on all Data in first 2 columns in first sheet
                    for (var row = 2; row < 9999; row++)
                    {
                        var changesMade = 0;
                        newFamilyTypeTransaction.Start();
                        if ((string)workSheet.Cells[row, 1].Value == null ||
                            (string)workSheet.Cells[row, 1].Value == "") break;
                        // add a new type and edit its parameters
                        var newFamilyType = familyManager.NewType((string)workSheet.Cells[row, 1].Value);
                        if (newFamilyType != null)
                        {
                            // look for 'b' and 'h' parameters and set them to 2 feet
                            var familyParam = familyManager.get_Parameter("b");
                            if (null != familyParam)
                            {
                                familyManager.Set(familyParam, MmToFoot((double)workSheet.Cells[row, 2].Value));
                                changesMade += 1;
                            }

                            familyParam = familyManager.get_Parameter("h");
                            if (null != familyParam)
                            {
                                familyManager.Set(familyParam, MmToFoot((double)workSheet.Cells[row, 3].Value));
                                changesMade += 1;
                            }

                            familyParam = familyManager.get_Parameter("Length");
                            if (null != familyParam) familyManager.Set(familyParam, MmToFoot(1524.0));
                        }

                        if (2 == changesMade) // set both paramaters?
                            newFamilyTypeTransaction.Commit();
                        else // could not make the change -> should roll back 
                            newFamilyTypeTransaction.RollBack();

                        // if could not make the change or could not commit it, we return
                        if (newFamilyTypeTransaction.GetStatus() != TransactionStatus.Committed)
                            TaskDialog.Show("Revit", "could not make the change or could not commit it");
                    }

                    // now update the Revit project with Family which has a new type
                    var loadOptions = new LoadOpts();

                    // This overload is necessary for reloading an edited family
                    // back into the source document from which it was extracted
                    family = familyDoc.LoadFamily(doc, loadOptions);
                    TaskDialog.Show("Revit", "New Beams Types Created Successfully");
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