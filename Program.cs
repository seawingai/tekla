using System;

namespace Tekla.ExcelMacros
{
    class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide an argument: export or import");
                return;
            }

            Console.WriteLine($"Starting {args[0].ToLower()} macro");

            switch (args[0].ToLower())
            {
                case "export":
                    var exporter = new ExportToExcel();
                    exporter.Export();
                    Console.WriteLine("Export complete.");
                    break;

                case "import":
                    var importer = new ImportFromExcel();
                    importer.Import();
                    Console.WriteLine("Import complete.");
                    break;

                default:
                    Console.WriteLine("Unknown argument. Use 'export' or 'import'.");
                    break;
            }

            Console.WriteLine($"Finished {args[0].ToLower()} macro");

            Console.ReadKey();
        }
    }
}
