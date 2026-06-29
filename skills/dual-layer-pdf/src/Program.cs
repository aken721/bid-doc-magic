using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace DualLayerPdfConverter
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length == 0 || args.Any(a => a == "--help" || a == "-h"))
            {
                PrintUsage();
                return args.Length == 0 ? 1 : 0;
            }

            var options = ParseArgs(args);
            if (options == null)
                return 1;

            try
            {
                PreloadNativeDlls();

                if (Directory.Exists(options.InputPath))
                    return BatchConvert(options);

                return SingleConvert(options);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"ERROR:{ex.Message}");
                return 2;
            }
        }

        static int SingleConvert(Options options)
        {
            var engine = new DualLayerPdfEngine
            {
                Dpi = options.Dpi,
                PdfInput = options.PdfInput,
                OpenAfter = options.OpenAfter,
                MaxDegreeOfParallelism = options.Threads
            };

            string outputPath = engine.Convert(options.InputPath, options.OutputPath);
            Console.WriteLine($"OK:{outputPath}");
            return 0;
        }

        static int BatchConvert(Options options)
        {
            var extensions = options.PdfInput
                ? new[] { ".pdf" }
                : new[] { ".docx", ".doc" };

            var files = extensions
                .SelectMany(ext => Directory.GetFiles(options.InputPath, $"*{ext}", SearchOption.TopDirectoryOnly))
                .Where(f => !Path.GetFileNameWithoutExtension(f).EndsWith("_DualPDF"))
                .ToArray();

            if (files.Length == 0)
            {
                Console.Error.WriteLine($"ERROR:No matching files found in: {options.InputPath}");
                return 2;
            }

            Console.WriteLine($"Found {files.Length} file(s) to convert.");

            int fileDop = options.Threads > 0 ? 1 : 1;
            int engineDop = options.Threads > 0 ? options.Threads : Environment.ProcessorCount;

            int success = 0;
            int failed = 0;
            var errors = new ConcurrentBag<string>();
            object lockObj = new object();

            var fileOptions = new ParallelOptions { MaxDegreeOfParallelism = fileDop };

            Parallel.ForEach(files, fileOptions, filePath =>
            {
                try
                {
                    string outputDir = !string.IsNullOrEmpty(options.OutputPath)
                        ? options.OutputPath
                        : Path.GetDirectoryName(filePath);

                    var engine = new DualLayerPdfEngine
                    {
                        Dpi = options.Dpi,
                        PdfInput = options.PdfInput,
                        OpenAfter = false,
                        MaxDegreeOfParallelism = engineDop
                    };

                    string outputPath = engine.Convert(filePath, Path.Combine(
                        outputDir,
                        Path.GetFileNameWithoutExtension(filePath) + "_DualPDF.pdf"));

                    lock (lockObj)
                    {
                        success++;
                        Console.WriteLine($"OK [{success + failed}/{files.Length}]: {outputPath}");
                    }
                }
                catch (Exception ex)
                {
                    lock (lockObj)
                    {
                        failed++;
                        errors.Add($"{filePath}: {ex.Message}");
                        Console.Error.WriteLine($"FAIL [{success + failed}/{files.Length}]: {filePath} - {ex.Message}");
                    }
                }
            });

            Console.WriteLine();
            Console.WriteLine($"Batch complete: {success} succeeded, {failed} failed out of {files.Length} file(s).");

            return failed == 0 ? 0 : 2;
        }

        static void PrintUsage()
        {
            Console.WriteLine("DualLayerPdfConverter - Convert DOCX/PDF to dual-layer PDF");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  Single file:  DualLayerPdfConverter -i <file> [options]");
            Console.WriteLine("  Batch dir:    DualLayerPdfConverter -i <folder> [options]");
            Console.WriteLine();
            Console.WriteLine("When -i points to a directory, all .docx/.doc files in it will be");
            Console.WriteLine("converted. Use --pdf-input to batch convert .pdf files instead.");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  -i, --input      Input file or directory path (required)");
            Console.WriteLine("  -o, --output     Output path: file path for single, directory for batch");
            Console.WriteLine("  -d, --dpi        Render DPI 50-1200 (default: 300)");
            Console.WriteLine("  -t, --threads    Max parallel threads per file (default: CPU core count)");
            Console.WriteLine("      --pdf-input  Treat input as PDF (skip Word-to-PDF step)");
            Console.WriteLine("      --open       Open output PDF after conversion (single file only)");
            Console.WriteLine("  -h, --help       Show this help message");
        }

        static Options ParseArgs(string[] args)
        {
            var opts = new Options();

            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "-i":
                    case "--input":
                        if (i + 1 >= args.Length) { Console.Error.WriteLine("ERROR:--input requires a value"); return null; }
                        opts.InputPath = args[++i];
                        break;
                    case "-o":
                    case "--output":
                        if (i + 1 >= args.Length) { Console.Error.WriteLine("ERROR:--output requires a value"); return null; }
                        opts.OutputPath = args[++i];
                        break;
                    case "-d":
                    case "--dpi":
                        if (i + 1 >= args.Length) { Console.Error.WriteLine("ERROR:--dpi requires a value"); return null; }
                        if (!int.TryParse(args[++i], out int dpi)) { Console.Error.WriteLine("ERROR:--dpi must be a number"); return null; }
                        opts.Dpi = dpi;
                        break;
                    case "-t":
                    case "--threads":
                        if (i + 1 >= args.Length) { Console.Error.WriteLine("ERROR:--threads requires a value"); return null; }
                        if (!int.TryParse(args[++i], out int threads) || threads < 1) { Console.Error.WriteLine("ERROR:--threads must be a positive number"); return null; }
                        opts.Threads = threads;
                        break;
                    case "--pdf-input":
                        opts.PdfInput = true;
                        break;
                    case "--open":
                        opts.OpenAfter = true;
                        break;
                    default:
                        Console.Error.WriteLine($"ERROR:Unknown option: {args[i]}");
                        return null;
                }
            }

            if (string.IsNullOrEmpty(opts.InputPath))
            {
                Console.Error.WriteLine("ERROR:--input is required");
                return null;
            }

            if (!File.Exists(opts.InputPath) && !Directory.Exists(opts.InputPath))
            {
                Console.Error.WriteLine($"ERROR:Input not found: {opts.InputPath}");
                return null;
            }

            if (opts.Dpi < 50 || opts.Dpi > 1200)
            {
                Console.Error.WriteLine("ERROR:--dpi must be between 50 and 1200");
                return null;
            }

            return opts;
        }

        static void PreloadNativeDlls()
        {
            string asmDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            if (string.IsNullOrEmpty(asmDir)) return;

            bool is64Bit = IntPtr.Size == 8;
            string srcName = is64Bit ? "pdfium_x64.dll" : "pdfium_x86.dll";
            string archSubDir = is64Bit ? "x64" : "x86";

            string[] candidates = new[]
            {
                Path.Combine(asmDir, srcName),
                Path.Combine(asmDir, archSubDir, "pdfium.dll"),
                Path.Combine(asmDir, "pdfium.dll"),
            };

            string srcPath = candidates.FirstOrDefault(p => File.Exists(p));
            if (srcPath == null) return;

            string dstPath = Path.Combine(asmDir, "pdfium.dll");

            if (srcPath != dstPath)
            {
                try
                {
                    bool needCopy = !File.Exists(dstPath);
                    if (!needCopy)
                    {
                        var srcInfo = new FileInfo(srcPath);
                        var dstInfo = new FileInfo(dstPath);
                        needCopy = srcInfo.Length != dstInfo.Length;
                    }
                    if (needCopy) File.Copy(srcPath, dstPath, true);
                }
                catch { }
            }

            try
            {
                SetDllDirectory(asmDir);
                string currentPath = Environment.GetEnvironmentVariable("PATH") ?? "";
                if (!currentPath.Contains(asmDir))
                    Environment.SetEnvironmentVariable("PATH", currentPath + ";" + asmDir);
            }
            catch { }

            try
            {
                string dllPath = File.Exists(dstPath) ? dstPath : srcPath;
                if (File.Exists(dllPath))
                    LoadLibrary(dllPath);
            }
            catch { }
        }

        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr LoadLibrary(string lpFileName);

        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);

        class Options
        {
            public string InputPath = "";
            public string OutputPath = "";
            public int Dpi = 300;
            public int Threads = 0;
            public bool PdfInput = false;
            public bool OpenAfter = false;
        }
    }
}
