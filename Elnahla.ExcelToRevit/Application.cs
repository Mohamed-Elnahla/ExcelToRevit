using Autodesk.Revit.UI;
using Elnahla.ExcelToRevit.Commands;
using Nice3point.Revit.Toolkit.External;
using Serilog.Events;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.Windows.Media;
using System.Windows.Forms;

namespace Elnahla.ExcelToRevit
{
    [UsedImplicitly]
    public class Application : ExternalApplication
    {
        public override void OnStartup()
        {
            CreateLogger();
            CreateRibbon();

        }

        public override void OnShutdown()
        {
            Log.CloseAndFlush();
        }

        private void CreateRibbon()
        {
            //Create Ribbon Tab
            Application.CreateRibbonTab("Elnahla");

            //Creating Panel
            var P_ExcelToRevit = Application.CreateRibbonPanel("Elnahla", "Excel To Revit");
            //var P_Levels = Application.CreateRibbonPanel("Elnahla", "Levels");
            //var P_Families = Application.CreateRibbonPanel("Elnahla", "Create Family Types");
            //var P_Place = Application.CreateRibbonPanel("Elnahla", "Place Elements");
            //var P_About = Application.CreateRibbonPanel("Elnahla", "About");

            //Creating Buttons
            var path = Assembly.GetExecutingAssembly().Location;
            var D_B_CLevels =
                new PushButtonData("B_CLevels", "Levels", path, "Elnahla.ExcelToRevit.Commands.CreatingLevels");
            var D_B_F_Piles = new PushButtonData("B_F_Piles", "Piles", path, "Elnahla.ExcelToRevit.Commands.PilesFamilies");
            var D_B_F_PC_Isolated_Footings = new PushButtonData("B_F_PC_Isolated_Footings", "Isolated PC Footings", path,
                "Elnahla.ExcelToRevit.Commands.F_PC_Isolated_Footings");
            var D_B_F_RC_Isolated_Footings = new PushButtonData("B_F_RC_Isolated_Footings", "Isolated RC Footings", path,
                "Elnahla.ExcelToRevit.Commands.F_RC_Isolated_Footings");
            var D_B_F_PC_RC_Slab_Footings = new PushButtonData("B_F_PC_RC_Slab_Footings", "Slab Foundation", path,
                "Elnahla.ExcelToRevit.Commands.SlabFoundations");
            var D_B_F_Columns =
                new PushButtonData("B_F_Columns", "Columns", path, "Elnahla.ExcelToRevit.Commands.ColumnsFamilies");
            var D_B_F_Walls = new PushButtonData("B_F_Walls", "Walls", path, "Elnahla.ExcelToRevit.Commands.WallsFamilies");
            var D_B_F_Beams = new PushButtonData("B_F_Beams", "Beams", path, "Elnahla.ExcelToRevit.Commands.BeamsFamilies");
            var D_B_F_StrFloors = new PushButtonData("B_F_StrFloors", "Structural Floors", path,
                "Elnahla.ExcelToRevit.Commands.StrFloorsFamilies");
            var D_B_F_TitleBlocks = new PushButtonData("B_F_TitleBlocks", "Titleblocks", path,
                "Elnahla.ExcelToRevit.Commands.TitleblocksFamilies");
            var D_Create_Family_Types = new PulldownButtonData("Create Family Types", "Create\r\nFamily Types");
            var D_B_Excel_Familes = new PushButtonData("B_Excel_Familes", "Families\r\nExcel Files", path,
                "Elnahla.ExcelToRevit.Commands.ExcelFamilies");
            var D_Place = new PulldownButtonData("Place Elements", "Place\r\nElements");
            var D_B_P_Piles =
                new PushButtonData("B_P_Piles", "Place Piles", path, "Elnahla.ExcelToRevit.Commands.PilesPlacing");
            var D_B_P_PC_Isolated_Footings = new PushButtonData("B_P_PC_Isolated_Footings", "Place Isolated PC Footings",
                path, "Elnahla.ExcelToRevit.Commands.P_PC_Isolated_Footings");
            var D_B_P_RC_Isolated_Footings = new PushButtonData("B_P_RC_Isolated_Footings", "Place Isolated RC Footings",
                path, "Elnahla.ExcelToRevit.Commands.P_RC_Isolated_Footings");
            var D_B_P_Columns =
                new PushButtonData("B_P_Columns", "Columns", path, "Elnahla.ExcelToRevit.Commands.ColumnsPlacing");
            var D_B_P_Walls = new PushButtonData("B_P_Walls", "Walls", path, "Elnahla.ExcelToRevit.Commands.WallsPlacing");
            var D_B_P_Beams = new PushButtonData("B_F_Beams", "Beams", path, "Elnahla.ExcelToRevit.Commands.BeamsPlacing");
            var D_B_Excel_Location = new PushButtonData("B_PExcel_Familes", "Placing Elemnets\r\nExcel Files", path,
                "Elnahla.ExcelToRevit.Commands.ExcelLocation");
            var D_B_About = new PushButtonData("B_About", "About", path, "Elnahla.ExcelToRevit.Commands.AboutCommand");


            //Create Image of Button;
            D_B_CLevels.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.Level.bmp");
            D_Create_Family_Types.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.Family.bmp");
            D_B_F_Piles.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.Piles.bmp");
            D_B_F_PC_Isolated_Footings.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.PC_Isolated_Footings.bmp");
            D_B_F_RC_Isolated_Footings.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.RC_Isolated_Footings.bmp");
            D_B_F_PC_RC_Slab_Footings.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.PC_Slab.bmp");
            D_B_F_Columns.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.C_Columns.bmp");
            D_B_F_Walls.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.Wall.bmp");
            D_B_F_Beams.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.Beams.bmp");
            D_B_F_StrFloors.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.Structural_Floor.bmp");
            D_B_F_TitleBlocks.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.Layout_Sheet.bmp");
            D_Place.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.Place.bmp");
            D_B_P_Piles.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.Piles.bmp");
            D_B_P_PC_Isolated_Footings.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.PC_Isolated_Footings.bmp");
            D_B_P_RC_Isolated_Footings.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.RC_Isolated_Footings.bmp");
            D_B_P_Columns.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.C_Columns.bmp");
            D_B_P_Walls.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.Wall.bmp");
            D_B_P_Beams.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.Beams.bmp");

            //stacked buttons
            D_B_Excel_Familes.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.Excel.bmp");
            D_B_Excel_Location.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.Excel.bmp");
            D_B_About.LargeImage = BmpImageSource("Elnahla.ExcelToRevit.Resources.About.bmp");
            //D_B_Excel_Familes.Image = BmpImageSource("Elnahla.ExcelToRevit.Resources.Excel.bmp");
            //D_B_Excel_Location.Image = BmpImageSource("Elnahla.ExcelToRevit.Resources.Excel.bmp");
            //D_B_About.Image = BmpImageSource("Elnahla.ExcelToRevit.Resources.About.bmp");

            //Place Buttons Into Panels
            //-------------------------
            //Create Levels Button
            var B_CLevels = P_ExcelToRevit.AddItem(D_B_CLevels) as PushButton;

            //Create Family Types Button
            var Create_Family_Types = P_ExcelToRevit.AddItem(D_Create_Family_Types) as PulldownButton;
            Create_Family_Types.AddPushButton(D_B_F_Piles);
            Create_Family_Types.AddPushButton(D_B_F_PC_Isolated_Footings);
            Create_Family_Types.AddPushButton(D_B_F_RC_Isolated_Footings);
            Create_Family_Types.AddPushButton(D_B_F_PC_RC_Slab_Footings);
            Create_Family_Types.AddPushButton(D_B_F_Columns);
            Create_Family_Types.AddPushButton(D_B_F_Walls);
            Create_Family_Types.AddPushButton(D_B_F_Beams);
            Create_Family_Types.AddPushButton(D_B_F_StrFloors);
            Create_Family_Types.AddPushButton(D_B_F_TitleBlocks);
            

            //Placing Elements Button
            var Place_Elements = P_ExcelToRevit.AddItem(D_Place) as PulldownButton;
            Place_Elements.AddPushButton(D_B_P_Piles);
            Place_Elements.AddPushButton(D_B_P_PC_Isolated_Footings);
            Place_Elements.AddPushButton(D_B_P_RC_Isolated_Footings);
            Place_Elements.AddPushButton(D_B_P_Columns);
            Place_Elements.AddPushButton(D_B_P_Beams);
            Place_Elements.AddPushButton(D_B_P_Walls);

            //Stacked Panel
            IList<RibbonItem> ribbonItem = P_ExcelToRevit.AddStackedItems(D_B_Excel_Familes, D_B_Excel_Location, D_B_About);
            //var B_Excel_Familes = P_ExcelToRevit.AddItem(D_B_Excel_Familes) as PushButton;
            //var B_Excel_Location = P_ExcelToRevit.AddItem(D_B_Excel_Location) as PushButton;
            //var B_About = P_ExcelToRevit.AddItem(D_B_About) as PushButton;
        }

        public static void CreateLogger()
        {
            const string outputTemplate = "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}";

            Log.Logger = new LoggerConfiguration()
                .WriteTo.Debug(LogEventLevel.Debug, outputTemplate)
                .MinimumLevel.Debug()
                .CreateLogger();

            AppDomain.CurrentDomain.UnhandledException += (_, args) =>
            {
                var e = (Exception)args.ExceptionObject;
                Log.Fatal(e, "Domain unhandled exception");
            };
        }

        private ImageSource BmpImageSource(string embeddedPath)
        {
            var stream = GetType().Assembly.GetManifestResourceStream(embeddedPath);
            var decoder = new BmpBitmapDecoder(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            return decoder.Frames[0];
        }
    }
}