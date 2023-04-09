using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Elnahla.ExcelToRevit.Commands;

[Transaction(TransactionMode.ReadOnly)]
internal class ExcelLocation : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        try
        {
            var path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var sourcePath = @"C:\Program Files\Elnahla Revit Plugin\Excel Files\Location";
            var targetPath = path + @"\Location Excel Files";

            // To copy a folder's contents to a new location:
            // Create a new target folder.
            // If the directory already exists, this method does not create a new directory.
            Directory.CreateDirectory(targetPath);

            // To copy all the files in one directory to another directory.
            // Get the files in the source folder. (To recursively iterate through
            // all subfolders under the current directory, see
            // "How to: Iterate Through a Directory Tree.")
            // Note: Check for target path was performed previously
            //       in this code example.
            if (Directory.Exists(sourcePath))
            {
                var files = Directory.GetFiles(sourcePath);

                // Copy the files and overwrite destination files if they already exist.
                foreach (var s in files)
                {
                    // Use static Path methods to extract only the file name from the path.
                    var fileName = Path.GetFileName(s);
                    var destFile = Path.Combine(targetPath, fileName);
                    File.Copy(s, destFile, true);
                }
            }
            else
            {
                Console.WriteLine("Source path does not exist!");
            }

            TaskDialog.Show("Revit", "Excel Files of Elements Locations Have Been Copied to Desktop");
            return Result.Succeeded;
        }
        catch (Exception e)
        {
            message = e.Message;
            return Result.Failed;
        }
    }
}