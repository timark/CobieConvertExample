using System;
using System.IO;
using System.Linq;
using Xbim.IO.CobieExpress;
using Xbim.Common;
using Xbim.CobieExpress;
using OfficeOpenXml;

namespace CobieConversionExample
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Loading Files");
            
            // Cobie file to process
            IModel model = CobieModel.ImportFromTable("E:\\Dormitory-ARC-AP.xlsx", out string report);
            
            // CAFM Template file
            var templateFileInfo = new FileInfo("E:\\Template.xlsx");
            ExcelPackage package = new ExcelPackage(templateFileInfo);

            // Results CAFM populated file
            var saveFileInfo = new FileInfo("E:\\TemplatePopulated.xlsx");

            // Set the workbook sheet to the Rooms sheet
            var ws = package.Workbook.Worksheets["Rooms"];
            var rowIndex = 0;

            // Get all Facility Details
            var facility = model.Instances.OfType<CobieFacility>().FirstOrDefault();
            Console.WriteLine($"Facility = {facility.Name}");
            // Get all Rooms in the project
            var rooms = model.Instances.OfType<CobieSpace>();
            Console.WriteLine($"Number of Rooms  = {rooms.Count()}");
            var r = 2; // start to write data to the 2nd row of the Rooms Sheet
            foreach (var room in rooms)
            {
                Console.WriteLine(room.Categories.FirstOrDefault().Value);
                ws.Cells[r, 1].Value = facility.Name;
                ws.Cells[r, 2].Value = room.Floor.Name;
                ws.Cells[r, 3].Value = room.Name;
                ws.Cells[r, 4].Value = room.Description;
                ws.Cells[r, 5].Value = room.Categories.FirstOrDefault().Value;
                ws.Cells[r, 6].Value = room.GrossArea;
                var wallFinish = room.Attributes.FirstOrDefault(x => x.Name == "Wall Finish");
                if (wallFinish != null)
                    ws.Cells[r, 7].Value = wallFinish.Value;
                var floorFinish = room.Attributes.FirstOrDefault(x => x.Name == "Floor Finish");
                if (floorFinish != null)
                    ws.Cells[r, 8].Value = floorFinish.Value;
                var ceilingFinish = room.Attributes.FirstOrDefault(x => x.Name == "Ceiling Finish");
                if (ceilingFinish != null)
                    ws.Cells[r, 9].Value = ceilingFinish.Value;

                r++;
                rowIndex++;
            }
            Console.WriteLine($"Exporting to File");

            package.SaveAs(saveFileInfo);

            package.Dispose(); // Not neccessary but tidy
            model.Dispose(); // Not neccessary but tidy

            Console.WriteLine($"Finished");
            Environment.Exit(0); // Not neccessary but tidy
        }

    }
}