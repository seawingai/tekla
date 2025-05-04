#pragma warning disable 1633 // Unrecognized #pragma directive
#pragma reference "Tekla.Macros.Akit"
#pragma reference "Tekla.Macros.Wpf.Runtime"
#pragma reference "Tekla.Macros.Runtime"
#pragma warning restore 1633 // Unrecognized #pragma directive

using System.Diagnostics;
using System.IO;
using Tekla.Structures.Model;

namespace UserMacros
{
    public sealed class Macro
    {
        [Tekla.Macros.Runtime.MacroEntryPointAttribute()]
        public static void Run(Tekla.Macros.Runtime.IMacroRuntime runtime)
        {
            var model = new Model();
            string modelPath = model.GetInfo().ModelPath;

            string exePath = Path.Combine(modelPath, @"macros\modeling\excel-macros\Debug\net48\ExcelMacros.exe");

            if (!File.Exists(exePath))
            {
                System.Windows.Forms.MessageBox.Show("Executable not found:\n" + exePath);
                return;
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = "import",
                UseShellExecute = true
            });
        }
    }
}
