using System;
using System.Collections;
using System.IO;
using ClosedXML.Excel;
using Tekla.Structures.Model;

namespace Tekla.ExcelMacros
{
    internal class ExportToExcel
    {
        public void Export()
        {
            try
            {
                Execute();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occured.\n{ex.Message}\n{ex.StackTrace}");

                Console.ReadKey();
            }
        }

        private void Execute()
        {
            var model = new Model();
            if (!model.GetConnectionStatus())
            {
                Console.WriteLine("Tekla not connected.");
                return;
            }

            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Model.xlsx");
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Members");

            string[] headers = { "ID", "Name", "Section", "Start_X", "Start_Y", "Start_Z", "End_X", "End_Y", "End_Z", "Comment" };
            for (int i = 0; i < headers.Length; i++)
                worksheet.Cell(1, i + 1).Value = headers[i];

            int row = 2;
            var selector = model.GetModelObjectSelector();
            var objects = selector.GetAllObjects();

            while (objects.MoveNext())
            {
                if (objects.Current is Beam beam)
                {
                    string comment = "";
                    beam.GetUserProperty("comment", ref comment);

                    worksheet.Cell(row, 1).Value = beam.Identifier.ID;
                    worksheet.Cell(row, 2).Value = beam.Name;
                    worksheet.Cell(row, 3).Value = beam.Profile.ProfileString;
                    worksheet.Cell(row, 4).Value = beam.StartPoint.X;
                    worksheet.Cell(row, 5).Value = beam.StartPoint.Y;
                    worksheet.Cell(row, 6).Value = beam.StartPoint.Z;
                    worksheet.Cell(row, 7).Value = beam.EndPoint.X;
                    worksheet.Cell(row, 8).Value = beam.EndPoint.Y;
                    worksheet.Cell(row, 9).Value = beam.EndPoint.Z;
                    worksheet.Cell(row, 10).Value = comment;

                    Console.WriteLine($"[{row}] - Beam added: {beam.Identifier.ID} - [{beam.Name}] - comment:[{comment}]");

                    row++;
                }
            }

            workbook.SaveAs(filePath);
            Console.WriteLine($"Excel saved: {filePath}");
        }

        private void GetUserPropertyNames(Beam beam)
        {
            Console.WriteLine($"UDA for beam: {beam.Name}");

            var enumProps = new Hashtable();
            beam.GetAllUserProperties(ref enumProps);

            foreach (DictionaryEntry entry in enumProps)
            {
                string prop = entry.Key.ToString();
                string value = entry.Value?.ToString() ?? "";
                Console.WriteLine(prop + " = " + value);
            }
        }
    }
}
