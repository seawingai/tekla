using System;
using System.IO;
using ClosedXML.Excel;
using Tekla.Structures;
using Tekla.Structures.Model;

namespace Tekla.ExcelMacros
{
    public class ImportFromExcel
    {
        public void Import()
        {
            var model = new Model();
            if (!model.GetConnectionStatus())
            {
                Console.WriteLine("Tekla not connected.");
                return;
            }

            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ExportedMembers.xlsx");
            if (!File.Exists(filePath))
            {
                Console.WriteLine("Excel file not found.");
                return;
            }

            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet("Members");

            int row = 2;
            while (!worksheet.Cell(row, 1).IsEmpty())
            {
                int id = Convert.ToInt32(worksheet.Cell(row, 1).Value);
                string comment = worksheet.Cell(row, 10).GetValue<string>();

                if (comment.Trim().Equals("Connect", StringComparison.OrdinalIgnoreCase))
                {
                    var obj = model.SelectModelObject(new Identifier(id));
                    if (obj is Beam beam)
                    {
                        Console.WriteLine($"Creating connection for Beam ID: {id}");

                        // Add connection creation logic here
                        // e.g., new Component { ... }
                    }
                }

                row++;
            }

            Console.WriteLine("Finished reading Excel and processing connections.");
        }
    }
}
