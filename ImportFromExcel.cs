using System;
using System.IO;
using System.Linq;
using Tekla.Structures.Model;
using ClosedXML.Excel;
using Tekla.Structures;

namespace Tekla.ExcelMacros
{
    public class ImportFromExcel
    {
        private readonly string excelFilePath;
        private readonly string sheetName = "Connections";
        private readonly string logFilePath;
        private Model model;

        public ImportFromExcel()
        {
            model = new Model();

            // Assuming the file is located on the Desktop
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            this.excelFilePath = Path.Combine(desktopPath, "Model.xlsx");
            this.logFilePath = Path.Combine(desktopPath, "ConnectionLog.txt");
        }

        public void Import()
        {
            if (!model.GetConnectionStatus())
            {
                Log("Cannot connect to Tekla Structures model.");
                return;
            }

            Log("=== Run started at " + DateTime.Now + " ===");

            if (!File.Exists(excelFilePath))
            {
                Log($"Excel file not found: {excelFilePath}");
                return;
            }

            try
            {
                var workbook = new XLWorkbook(excelFilePath);

                if (!workbook.Worksheets.Contains(sheetName))
                {
                    Log($"Sheet '{sheetName}' not found in Excel file.");
                    return;
                }

                var ws = workbook.Worksheet(sheetName);
                var rows = ws.RangeUsed().RowsUsed().Skip(1)
                    .Where(r => !string.IsNullOrWhiteSpace(r.Cell("J").GetString()))
                    .ToList();

                foreach (var row in rows)
                {
                    try
                    {
                        int primaryId = int.Parse(row.Cell("A").GetString());
                        int secondaryId = int.Parse(row.Cell("E").GetString());
                        int connectionType = int.Parse(row.Cell("J").GetString());

                        var primary = model.SelectModelObject(new Identifier(primaryId)) as Part;
                        var secondary = model.SelectModelObject(new Identifier(secondaryId)) as Part;

                        if (primary == null || secondary == null)
                        {
                            Log($"Parts not found for IDs {primaryId}, {secondaryId}");
                            continue;
                        }

                        if (ConnectionExists(primary, secondary, connectionType))
                        {
                            Log($"Connection already exists between {primaryId} and {secondaryId}");
                            continue;
                        }

                        string connectionName = GetConnectionName(connectionType);

                        var connection = new Connection
                        {
                            Name = connectionName,
                            Number = connectionType
                        };

                        connection.SetPrimaryObject(primary);
                        connection.SetSecondaryObject(secondary);

                        if (!connection.Insert())
                        {
                            Log($"Failed to insert {connectionType} between {primaryId} and {secondaryId}");
                        }
                        else
                        {
                            Log($"Inserted {connectionType} between {primaryId} and {secondaryId}");
                        }
                    }
                    catch (Exception exRow)
                    {
                        Log($"Row {row.RowNumber()} error: {exRow.Message}");
                    }
                }

                model.CommitChanges();
            }
            catch (Exception ex)
            {
                Log("Fatal Error: " + ex.ToString());
            }

            Log("=== Run ended at " + DateTime.Now + " ===");
        }

        private bool ConnectionExists(Part primary, Part secondary, int number)
        {
            var children = primary.GetChildren();
            while (children.MoveNext())
            {
                if (children.Current is Connection connection && connection.Number == number)
                {
                    foreach (var obj in connection.GetSecondaryObjects())
                    {
                        if (obj is Part secPart && secPart.Identifier.ID == secondary.Identifier.ID)
                            return true;
                    }
                }
            }
            return false;
        }

        private string GetConnectionName(int connectionType)
        {
            return connectionType switch
            {
                77 => "ColumnSplice",           // or "BeamSplice", "BracingSplice"
                40 => "MomentEndPlate",
                14 => "PinEndPlateFlange",      // also used for some brace connections
                119 => "PinEndPlateWeb",
                0 => "PinWebPlate",
                27 => "BeamToBeamPinEndPlate1",
                65 => "BeamToBeamPinEndPlate2",
                185 => "SingleOuterPlate",
                146 => "ShearPlate",
                11 => "SingleAngle",            // or could be 10
                10 => "SingleAngle",            // in case 10 is used
                20 => "EndPlate",
                999 => "Unknown",
                _ => "Unknown"
            };
        }

        private void Log(string message)
        {
            string fullMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
            Console.WriteLine(fullMessage);
            File.AppendAllText(logFilePath, fullMessage + Environment.NewLine);
        }
    }
}
