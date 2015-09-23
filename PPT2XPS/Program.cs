using System;
using System.Linq;
using System.Reflection;
using System.IO;
using System.Diagnostics;

namespace PPT2XPS
{
    class Program
    {
        static void Main(string[] args)
        {
            // Use path from 1st argument or exe location.
            string path = GetParamOrExePath(args.FirstOrDefault());

            PrintAll(path);

            Console.WriteLine("");
            Console.WriteLine("Press any key to continue . . .");
            Console.ReadKey(true);
        }

        private static string GetParamOrExePath(string paramPath)
        {
            return paramPath?.ToString() ?? Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        }

        private static void PrintAll(string path)
        {
            Stopwatch watch = Stopwatch.StartNew();

            string[] files = GetFiles(path, "*.pptx");

            if (files.Length > 0)
            {
                PrintFiles(files);

                Console.WriteLine("Done converting PPTX to XPS!");

                // Time taken.
                watch.Stop();
                TimeSpan ts = watch.Elapsed;
                Console.WriteLine($"Total time taken: {ts.Minutes:00}:{ts.Seconds:00}");
            }
            else
            {
                watch.Stop();
                Console.WriteLine("No files to convert.");
            }
        }

        private static string[] GetFiles(string path, string searchPattern)
        {
            Console.WriteLine($"Listing {searchPattern} files in {path} ...");
            Console.WriteLine("");
            string[] files = Directory.GetFiles(path, searchPattern);

            return files;
        }

        private static void PrintFiles(string[] files)
        {
            for (int i = 0; i < files.Length; i++)
            {
                Stopwatch watch = Stopwatch.StartNew();

                try
                {
                    // Print file.
                    Ppt2XpsConverter.PrintFileAsXps(files[i], i + 1, files.Length);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    return;
                }

                // Time taken.
                watch.Stop();
                TimeSpan ts = watch.Elapsed;
                Console.WriteLine($"Time taken: {ts.Minutes:00}:{ts.Seconds:00}");
                Console.WriteLine("");
            }
        }
    }
}
