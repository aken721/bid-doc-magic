using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text.pdf;
using PDFiumSharp;
using PDFiumSharp.Types;
using Word = Microsoft.Office.Interop.Word;

namespace DualLayerPdfConverter
{
    public class DualLayerPdfEngine
    {
        public int Dpi { get; set; } = 300;
        public bool PdfInput { get; set; } = false;
        public bool OpenAfter { get; set; } = false;
        public int MaxDegreeOfParallelism { get; set; } = 0;

        private static readonly string TempDir = Path.Combine(Path.GetTempPath(), "BidDocMagic");

        public string Convert(string inputPath, string outputPath = null)
        {
            inputPath = Path.GetFullPath(inputPath);

            string sessionDir = Path.Combine(TempDir, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(sessionDir);

            int dop = MaxDegreeOfParallelism > 0 ? MaxDegreeOfParallelism : Environment.ProcessorCount;

            try
            {
                string textPdfPath;

                if (PdfInput)
                {
                    string srcPdf = EnsureAsciiPath(inputPath);
                    textPdfPath = Path.Combine(sessionDir, "text_layer.pdf");
                    File.Copy(srcPdf, textPdfPath, true);
                }
                else
                {
                    Console.WriteLine("[1/3] Converting DOCX to PDF text layer...");
                    textPdfPath = Path.Combine(sessionDir, "text_layer.pdf");
                    ConvertWordToPDF(inputPath, textPdfPath);
                }

                string imagePattern = Path.Combine(sessionDir, "page-{0}.png");

                Console.WriteLine($"[2/3] Rendering with PDFium ({Dpi}dpi)...");
                RenderPdfToImages(textPdfPath, imagePattern, Dpi);

                if (string.IsNullOrEmpty(outputPath))
                {
                    string dir = Path.GetDirectoryName(inputPath);
                    string name = Path.GetFileNameWithoutExtension(inputPath);
                    outputPath = ResolveOutputPath(dir, name);
                }
                else
                {
                    outputPath = ResolveOutputPath(outputPath);
                }

                string finalOutputPath = outputPath;
                bool needMove = outputPath.Any(c => c > 127);
                string tempOutputPath = needMove
                    ? Path.Combine(sessionDir, "output_DualPDF.pdf")
                    : outputPath;

                Console.WriteLine($"[3/3] Composing dual-layer PDF (parallel={Math.Min(dop, 4)})...");
                CreatePdfWithImageOverlay(textPdfPath, imagePattern, tempOutputPath, false, sessionDir);

                if (needMove && File.Exists(tempOutputPath))
                {
                    File.Copy(tempOutputPath, finalOutputPath, true);
                }

                Console.WriteLine($"Done: {finalOutputPath}");

                if (OpenAfter && File.Exists(finalOutputPath))
                    Process.Start(new ProcessStartInfo(finalOutputPath) { UseShellExecute = true });

                return finalOutputPath;
            }
            finally
            {
                try
                {
                    if (Directory.Exists(sessionDir))
                        Directory.Delete(sessionDir, true);
                }
                catch { }

                try
                {
                    if (Directory.Exists(TempDir) && !Directory.EnumerateFileSystemEntries(TempDir).Any())
                        Directory.Delete(TempDir, false);
                }
                catch { }
            }
        }

        private void ConvertWordToPDF(string docxPath, string pdfPath)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application();
                wordApp.Visible = false;
                wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

                doc = wordApp.Documents.Open(docxPath, ReadOnly: true, AddToRecentFiles: false, Visible: false);
                doc.SaveAs2(
                    FileName: Path.GetFullPath(pdfPath),
                    FileFormat: Word.WdSaveFormat.wdFormatPDF,
                    AddToRecentFiles: false,
                    EmbedTrueTypeFonts: true
                );
            }
            finally
            {
                try
                {
                    if (doc != null)
                    {
                        doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                        Marshal.ReleaseComObject(doc);
                    }
                }
                catch { }

                try
                {
                    if (wordApp != null)
                    {
                        wordApp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                        Marshal.ReleaseComObject(wordApp);
                    }
                }
                catch { }
            }
        }

        private void RenderPdfToImages(string pdfPath, string imagePattern, int dpi)
        {
            using (var doc = new PDFiumSharp.PdfDocument(pdfPath))
            {
                int totalPages = doc.Pages.Count;

                for (int i = 0; i < totalPages; i++)
                {
                    var page = doc.Pages[i];
                    float scale = dpi / 72f;
                    int width = (int)(page.Width * scale);
                    int height = (int)(page.Height * scale);

                    using (var pdfBitmap = new PDFiumBitmap(width, height, false))
                    {
                        pdfBitmap.Fill(0xFFFFFFFF);
                        page.Render(pdfBitmap);

                        using (var bitmap = new Bitmap(width, height, pdfBitmap.Stride,
                            PixelFormat.Format32bppRgb, pdfBitmap.Scan0))
                        {
                            string imagePath = string.Format(imagePattern, i + 1);
                            bitmap.Save(imagePath, ImageFormat.Png);
                        }
                    }

                    Console.WriteLine($"  Rendering page {i + 1}/{totalPages}...");
                }
            }
        }

        private void CreatePdfWithImageOverlay(string textPdfPath, string imagePattern,
            string outputPath, bool isBackGround, string sessionDir)
        {
            string pageDir = Path.Combine(sessionDir, "pages");
            Directory.CreateDirectory(pageDir);

            int totalPages;
            var pageSizes = new Dictionary<int, iTextSharp.text.Rectangle>();

            ArrayList bookmarks = null;

            var mainReader = new PdfReader(textPdfPath);
            try
            {
                totalPages = mainReader.NumberOfPages;
                for (int i = 1; i <= totalPages; i++)
                    pageSizes[i] = mainReader.GetPageSize(i);
                bookmarks = SimpleBookmark.GetBookmark(mainReader);
            }
            finally
            {
                mainReader.Close();
            }

            Console.WriteLine($"  Splitting text PDF into {totalPages} pages...");
            for (int i = 1; i <= totalPages; i++)
            {
                var pageReader = new PdfReader(textPdfPath);
                try
                {
                    pageReader.SelectPages(i.ToString());
                    var fs = new FileStream(Path.Combine(pageDir, $"text-{i}.pdf"), FileMode.Create);
                    var stamper = new PdfStamper(pageReader, fs);
                    stamper.Close();
                    pageReader.Close();
                    fs.Close();
                }
                catch
                {
                    try { pageReader.Close(); } catch { }
                }
            }

            int composeDop = Math.Min(MaxDegreeOfParallelism > 0 ? MaxDegreeOfParallelism : Environment.ProcessorCount, 4);
            Console.WriteLine($"  Composing {totalPages} pages (parallel={composeDop})...");

            int completed = 0;
            object lockObj = new object();
            var composeOptions = new ParallelOptions { MaxDegreeOfParallelism = composeDop };

            Parallel.For(1, totalPages + 1, composeOptions, pageNum =>
            {
                string singlePagePath = Path.Combine(pageDir, $"text-{pageNum}.pdf");
                string overlayPath = Path.Combine(pageDir, $"dual-{pageNum}.pdf");
                string imagePath = string.Format(imagePattern, pageNum);

                if (!File.Exists(imagePath))
                    throw new FileNotFoundException($"Page {pageNum} image not found: {imagePath}");

                var pageReader = new PdfReader(singlePagePath);
                FileStream fs = null;
                PdfStamper stamper = null;

                try
                {
                    fs = new FileStream(overlayPath, FileMode.Create);
                    stamper = new PdfStamper(pageReader, fs);

                    iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(imagePath);
                    image.SetAbsolutePosition(0, 0);
                    image.ScaleAbsolute(pageSizes[pageNum].Width, pageSizes[pageNum].Height);

                    PdfContentByte canvas = isBackGround
                        ? stamper.GetUnderContent(1)
                        : stamper.GetOverContent(1);
                    canvas.AddImage(image);
                }
                finally
                {
                    try { stamper?.Close(); } catch { }
                    try { pageReader.Close(); } catch { }
                    try { fs?.Close(); } catch { }
                }

                lock (lockObj)
                {
                    completed++;
                    if (completed % 5 == 0 || completed == totalPages)
                        Console.WriteLine($"  Composing page {completed}/{totalPages}...");
                }
            });


            Console.WriteLine($"  Merging {totalPages} pages into final PDF...");
            MergePdfs(pageDir, "dual-", totalPages, outputPath, bookmarks);
        }

        private void MergePdfs(string pageDir, string prefix, int totalPages, string outputPath,
            ArrayList bookmarks = null)
        {
            if (string.IsNullOrEmpty(outputPath))
                throw new ArgumentException("Output path cannot be empty.", nameof(outputPath));

            var document = new iTextSharp.text.Document();
            var copy = new PdfCopy(document, new FileStream(outputPath, FileMode.Create));
            document.Open();

            try
            {
                for (int i = 1; i <= totalPages; i++)
                {
                    string pagePath = Path.Combine(pageDir, $"{prefix}{i}.pdf");
                    if (!File.Exists(pagePath))
                        throw new FileNotFoundException($"Merged page file not found: {pagePath}");

                    var reader = new PdfReader(pagePath);
                    try
                    {
                        for (int p = 1; p <= reader.NumberOfPages; p++)
                        {
                            var imported = copy.GetImportedPage(reader, p);
                            copy.AddPage(imported);
                        }
                    }
                    finally
                    {
                        try { reader.Close(); } catch { }
                    }
                }

                if (bookmarks != null && bookmarks.Count > 0)
                    copy.Outlines = bookmarks;
            }
            finally
            {
                document.Close();
                copy.Close();
            }
        }

        private string ResolveOutputPath(string directory, string baseName)
        {
            string candidate = Path.Combine(directory, baseName + "_DualPDF.pdf");
            if (!File.Exists(candidate))
                return candidate;

            for (int i = 1; i < 100; i++)
            {
                candidate = Path.Combine(directory, $"{baseName}_DualPDF({i:D2}).pdf");
                if (!File.Exists(candidate))
                    return candidate;
            }

            int seq = 100;
            while (true)
            {
                candidate = Path.Combine(directory, $"{baseName}_DualPDF({seq}).pdf");
                if (!File.Exists(candidate))
                    return candidate;
                seq++;
            }
        }

        private string ResolveOutputPath(string outputPath)
        {
            if (!File.Exists(outputPath))
                return outputPath;

            string dir = Path.GetDirectoryName(outputPath);
            string name = Path.GetFileNameWithoutExtension(outputPath);

            for (int i = 1; i < 100; i++)
            {
                string candidate = Path.Combine(dir, $"{name}({i:D2}).pdf");
                if (!File.Exists(candidate))
                    return candidate;
            }

            int seq = 100;
            while (true)
            {
                string candidate = Path.Combine(dir, $"{name}({seq}).pdf");
                if (!File.Exists(candidate))
                    return candidate;
                seq++;
            }
        }

        private string EnsureAsciiPath(string filePath)
        {
            bool hasNonAscii = filePath.Any(c => c > 127);
            if (!hasNonAscii) return filePath;

            string asciiPath = Path.Combine(TempDir, "input_temp.pdf");
            File.Copy(filePath, asciiPath, true);
            return asciiPath;
        }
    }
}
