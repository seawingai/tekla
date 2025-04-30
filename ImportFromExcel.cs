using System;
using System.IO;
using System.Linq;
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
                Log("Error: Tekla model is not connected.");
                return;
            }

            string modelPath = model.GetInfo().ModelPath;
            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Connections.xlsx");

            if (!File.Exists(filePath))
            {
                Log($"Error: Excel file not found: {filePath}");
                return;
            }

            var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheets.First();

            var validRows = worksheet.RowsUsed().Skip(1) // Skip header
                .Where(row => !string.IsNullOrWhiteSpace(row.Cell(5).GetString())) // 'Selected' column
                .ToList();

            Log($"Starting import. {validRows.Count} connections to process.");

            foreach (var row in validRows)
            {
                try
                {
                    string guid1 = row.Cell(1).GetString();
                    string guid2 = row.Cell(2).GetString();
                    string connectionType = row.Cell(3).GetString();
                    string attributeFile = row.Cell(4).GetString();

                    var part1 = model.SelectModelObject(new Identifier(new Guid(guid1))) as Part;
                    var part2 = model.SelectModelObject(new Identifier(new Guid(guid2))) as Part;

                    if (part1 == null || part2 == null)
                    {
                        Log($"Warning: One or both parts not found. Skipping: {guid1}, {guid2}");
                        continue;
                    }

                    int connectionNumber = GetConnectionNumber(connectionType);
                    if (connectionNumber == -1)
                    {
                        Log($"Warning: Unknown connection type '{connectionType}'. Skipping.");
                        continue;
                    }

                    if (ConnectionExists(part1, part2, connectionNumber))
                    {
                        Log($"Info: Connection already exists between {guid1} and {guid2}. Skipping.");
                        continue;
                    }

                    var connection = new Connection
                    {
                        Name = connectionType,
                        Number = connectionNumber
                    };

                    connection.SetPrimaryObject(part1);
                    connection.SetSecondaryObject(part2);

                    if (!string.IsNullOrWhiteSpace(attributeFile))
                        connection.LoadAttributesFromFile(attributeFile);
                    else
                        connection.LoadAttributesFromFile("standard");

                    if (connection.Insert() {
                        Log($"Success: Inserted {connectionType} between {guid1} and {guid2}.");
                    }
                    else {
                        Log($"Error: Failed to insert {connectionType} between {guid1} and {guid2}.");
                    }
                }
                catch (Exception ex)
                {
                    Log($"Exception: {ex.Message}");
                }
            }

            model.CommitChanges();
            Log("Import complete.");
        }

        private int GetConnectionNumber(string type)
        {
            switch (type.ToLowerInvariant())
            {
                case "endplate": return 144;
                case "clipangle": return 34;
                case "baseplate": return 1001;
                case "momentconnection": return 103;
                case "splice": return 1003;
                case "tubeconnection": return 116;
                case "gussetplate": return 40;
                case "haunch": return 164;
                case "stiffenedendplate": return 102;
                default: return -1;
            }
        }
        private bool ConnectionExists(Part primary, Part secondary, int number)
        {
            ModelObjectEnumerator children = primary.GetChildren();

            while (children.MoveNext())
            {
                if (children.Current is Connection connection)
                {
                    if (connection.Number == number)
                    {
                        var secondaries = connection.GetSecondaryObjects(); 

                        foreach (var obj in secondaries)
                        {
                            if (obj is Part secPart && secPart.Identifier.ID == secondary.Identifier.ID)
                            {
                                return true;
                            }
                        }
                    }
                }
            }

            return false;
        }

        private void Log(string message)
        {
            string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
            Console.WriteLine(line);
        }
    }
}
