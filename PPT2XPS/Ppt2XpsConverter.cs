using System;
using System.IO;
using System.Runtime.InteropServices;
using MSOInterop = Microsoft.Office.Interop;
using MSOCore = Microsoft.Office.Core;

namespace PPT2XPS
{
    public static class Ppt2XpsConverter
    {
        public static void PrintFileAsXps(string path, int index, int total)
        {
            // Avoid Quit if already running.
            bool appAlreadyRunning = System.Diagnostics.Process.GetProcessesByName("powerpnt").Length > 0;

            // Create/get instance of MS PowerPoint.
            // Do it once per file to avoid out-of-memory issues or crashes to due memory fragmentation, etc.
            Console.WriteLine("Starting MS PowerPoint ...");
            var app = new MSOInterop.PowerPoint.Application();

            // Open presentation.
            MSOInterop.PowerPoint._Presentation pptFile;

            var fileInfo = new FileInfo(path);
            Console.WriteLine($"Opening {Path.GetFileName(path)} ({(double)fileInfo.Length / 1024 / 1024:0.##} MB) {index}/{total} ...");
            pptFile = app.Presentations.Open(path,
                                             WithWindow: MSOCore.MsoTriState.msoFalse); // Don't show window.

            // Set options.
            SetPrintOptions(pptFile);

            // Delete if exists.
            DeleteIfExists(pptFile);

            // Print.
            Console.WriteLine($"Printing ...");
            pptFile.PrintOut(PrintToFile: pptFile.FullName.Replace(".pptx", ".xps"));

            // Close file.
            pptFile.Close();
            Marshal.ReleaseComObject(pptFile);
            pptFile = null;

            // Quit app.
            if (!appAlreadyRunning)
            {
                app.Quit();
            }

            // Force process to exit.
            Marshal.ReleaseComObject(app);
            app = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private static void DeleteIfExists(MSOInterop.PowerPoint._Presentation pptFile)
        {
            string xpsPath = pptFile.FullName.Replace(".pptx", ".xps");
            if (File.Exists(xpsPath))
            {
                Console.WriteLine($"Deleting existing file ...");
                File.Delete(xpsPath);
            }
        }

        private static void SetPrintOptions(MSOInterop.PowerPoint._Presentation pptFile)
        {
            pptFile.PrintOptions.ActivePrinter = "Microsoft XPS Document Writer";
            pptFile.PrintOptions.OutputType = MSOInterop.PowerPoint.PpPrintOutputType.ppPrintOutputNotesPages;
            pptFile.PrintOptions.PrintColorType = MSOInterop.PowerPoint.PpPrintColorType.ppPrintColor;
            pptFile.PrintOptions.HighQuality = MSOCore.MsoTriState.msoFalse; // msoTrue = -1 | msoFalse = 0
            pptFile.PrintOptions.PrintHiddenSlides = MSOCore.MsoTriState.msoTrue;
            pptFile.PrintOptions.PrintInBackground = MSOCore.MsoTriState.msoFalse; // Wait for print to finish.
        }
    }
}
