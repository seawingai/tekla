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
        private Model model;

        public ImportFromExcel()
        {
            model = new Model();

            // Assuming the file is located on the Desktop
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            this.excelFilePath = Path.Combine(desktopPath, "Model.xlsx");
        }

        public void Import()
        {
            try
            {
                Execute();
            }
            catch (Exception ex)
            {
                Log($"Fatal Error occured.\n{ex.Message}\n{ex.StackTrace}");

                Console.ReadKey();
            }
        }

        private void Execute()
        {
            Log($"Import started");

            if (!model.GetConnectionStatus())
            {
                Log("Cannot connect to Tekla Structures model.");
                return;
            }

            if (!File.Exists(excelFilePath))
            {
                Log($"Excel file not found: {excelFilePath}");
                return;
            }

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
                    /*
                        A = Preliminary ID  
                        B = Pre. Name  
                        C = Pre. Section  
                        D = Pre. Comment  
                        E = Secondary ID  
                        F = Sec. Name  
                        G = Sec. Section  
                        H = Sec. Comment  
                        I = Mid-Connection  
                        J = Connection Type  
                    */
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

                    var connection = new Connection
                    {
                        Number = connectionType
                    };

                    connection.SetPrimaryObject(primary);
                    connection.SetSecondaryObject(secondary);

                    if (!connection.Insert())
                    {
                        Log($"Failed to insert {connectionType} between {primaryId} and {secondaryId}. May be connection already exists.");
                    }
                    else
                    {
                        Log($"Successfully inserted {connectionType} between {primaryId} and {secondaryId}");
                    }
                }
                catch (Exception exRow)
                {
                    Log($"Row {row.RowNumber()} error: {exRow.Message}");
                }
            }

            model.CommitChanges();

            Log($"Import finished");
        }

        private void Log(string message)
        {
            string fullMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
            Console.WriteLine(fullMessage);
        }
    }
}
