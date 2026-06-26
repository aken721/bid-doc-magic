using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using iTextSharp.text;
using iTextSharp.text.pdf;
using PDFiumSharp;
using PDFiumSharp.Types;

namespace BidDocMagic
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        #region 双层PDF功能

        private static int _pdfDpi = 300;
        private static bool _pdfOpenAfter = false;

        private static readonly string TempDir = Path.Combine(Path.GetTempPath(), "BidDocMagic");

        private void btnPdfSettings_Click(object sender, RibbonControlEventArgs e)
        {
            using (var dialog = new PdfSettingsDialog())
            {
                dialog.InitialDpi = _pdfDpi;
                dialog.InitialOpenAfterConvert = _pdfOpenAfter;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _pdfDpi = dialog.Dpi;
                    _pdfOpenAfter = dialog.OpenAfterConvert;
                }
            }
        }

        private ProgressForm _progressForm;
        private System.Threading.CancellationTokenSource _cts;

        private void btnConvertToPDF_Click(object sender, RibbonControlEventArgs e)
        {
            btnConvertToPDF.Enabled = false;
            _cts = new System.Threading.CancellationTokenSource();

            _progressForm = new ProgressForm();
            _progressForm.CancellationTokenSource = _cts;
            _progressForm.FormClosed += (s, ea) =>
            {
                btnConvertToPDF.Enabled = true;
                _progressForm = null;
            };

            string docPath = ThisAddIn.app.ActiveDocument.FullName;
            string docName = ThisAddIn.app.ActiveDocument.Name;

            var thread = new System.Threading.Thread(() =>
            {
                string outputPath = null;
                bool success = false;
                string errorMsg = null;

                try
                {
                    string sessionDir = Path.Combine(TempDir, Guid.NewGuid().ToString("N"));
                    Directory.CreateDirectory(sessionDir);

                    try
                    {
                        if (_cts.Token.IsCancellationRequested) return;

                        UpdateProgress("步骤1/3：正在导出PDF文本层...");

                        string textPdfPath = Path.Combine(sessionDir, "text_layer.pdf");
                        ConvertWordToPDF(textPdfPath);

                        if (_cts.Token.IsCancellationRequested) return;

                        string imagePattern = Path.Combine(sessionDir, "page-{0}.png");
                        int dop = Environment.ProcessorCount;

                        UpdateProgress($"步骤2/3：正在使用PDFium渲染（{_pdfDpi}dpi）...");
                        RenderPdfToImages(textPdfPath, imagePattern, _pdfDpi);

                        if (_cts.Token.IsCancellationRequested) return;

                        string outputDir = Path.GetDirectoryName(docPath);
                        string baseName = Path.GetFileNameWithoutExtension(docName);
                        outputPath = ResolveOutputPath(outputDir, baseName);

                        string tempOutputPath = outputPath;
                        bool needMoveOutput = outputPath.Any(c => c > 127);
                        if (needMoveOutput)
                        {
                            tempOutputPath = Path.Combine(sessionDir, "output_DualPDF.pdf");
                        }

                        UpdateProgress($"步骤3/3：正在合成双层PDF（并行={dop}）...");
                        CreatePdfWithImageOverlayParallel(textPdfPath, imagePattern, tempOutputPath, false, dop, sessionDir);

                        if (_cts.Token.IsCancellationRequested) return;

                        if (needMoveOutput && File.Exists(tempOutputPath))
                        {
                            File.Copy(tempOutputPath, outputPath, true);
                        }

                        success = true;
                    }
                    finally
                    {
                        try
                        {
                            if (Directory.Exists(sessionDir))
                                Directory.Delete(sessionDir, true);
                        }
                        catch { }
                    }
                }
                catch (OperationCanceledException) { }
                catch (AggregateException aex)
                {
                    aex.Handle(ex =>
                    {
                        if (ex is OperationCanceledException) return true;
                        errorMsg = ex.Message;
                        return true;
                    });
                }
                catch (Exception ex)
                {
                    errorMsg = ex.Message;
                }

                if (success && outputPath != null)
                {
                    if (_pdfOpenAfter)
                    {
                        _progressForm?.CloseProgress();
                        try { Process.Start(new ProcessStartInfo(outputPath) { UseShellExecute = true }); } catch { }
                    }
                    else
                    {
                        _progressForm?.ShowResult("转换成功！输出文件：" + outputPath, outputPath);
                    }
                }
                else if (errorMsg != null)
                {
                    _progressForm?.ShowError("转换失败：" + errorMsg);
                }
                else
                {
                    _progressForm?.CloseProgress();
                }
            });
            thread.IsBackground = true;
            thread.SetApartmentState(System.Threading.ApartmentState.STA);

            _progressForm.Shown += (s, ea) =>
            {
                _progressForm.BeginInvoke(new Action(() => thread.Start()));
            };

            _progressForm.ShowDialog();
        }

        private void UpdateProgress(string text)
        {
            _progressForm?.UpdateStatus(text);
        }


        private void SetStatus(string text)
        {
            try
            {
                ThisAddIn.app.StatusBar = text;
            }
            catch { }
        }

        private void ConvertWordToPDF(string pdfFilePath)
        {
            Word.Document doc = ThisAddIn.app.ActiveDocument;
            doc.SaveAs2(
                FileName: Path.GetFullPath(pdfFilePath),
                FileFormat: Word.WdSaveFormat.wdFormatPDF,
                AddToRecentFiles: false,
                EmbedTrueTypeFonts: true,
                SaveFormsData: false,
                SaveAsAOCELetter: false,
                Encoding: Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8
            );
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

                        using (var bitmap = new Bitmap(width, height, pdfBitmap.Stride, PixelFormat.Format32bppRgb, pdfBitmap.Scan0))
                        {
                            string imagePath = string.Format(imagePattern, i + 1);
                            bitmap.Save(imagePath, ImageFormat.Png);
                        }
                    }

                    UpdateProgress($"步骤2/3：正在渲染第 {i + 1}/{totalPages} 页（PDFium {_pdfDpi}dpi）...");
                }
            }
        }

        private static readonly object _itextLock = new object();

        private void CreatePdfWithImageOverlayParallel(string textPdfPath, string imagePattern, string outputPath, bool isBackGround, int dop, string sessionDir)
        {
            var token = _cts?.Token ?? System.Threading.CancellationToken.None;

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

            UpdateProgress($"步骤3/3：正在拆分PDF为{totalPages}个单页...");

            for (int i = 1; i <= totalPages; i++)
            {
                token.ThrowIfCancellationRequested();
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

            int composeDop = Math.Min(dop, 4);
            UpdateProgress($"步骤3/3：正在并行合成{totalPages}页（并行={composeDop}）...");

            int completed = 0;
            object lockObj = new object();
            var composeOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = composeDop
            };

            Parallel.For(1, totalPages + 1, composeOptions, pageNum =>
            {
                token.ThrowIfCancellationRequested();

                string singlePagePath = Path.Combine(pageDir, $"text-{pageNum}.pdf");
                string overlayPath = Path.Combine(pageDir, $"dual-{pageNum}.pdf");
                string imagePath = string.Format(imagePattern, pageNum);

                if (!File.Exists(imagePath))
                    throw new FileNotFoundException($"页码 {pageNum} 的图片未找到: {imagePath}");

                var pageReader = new PdfReader(singlePagePath);
                FileStream fs = null;
                PdfStamper stamper = null;

                try
                {
                    fs = new FileStream(overlayPath, FileMode.Create);
                    stamper = new PdfStamper(pageReader, fs);

                    token.ThrowIfCancellationRequested();

                    iTextSharp.text.Image image;
                    lock (_itextLock)
                    {
                        image = iTextSharp.text.Image.GetInstance(imagePath);
                    }
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
                        UpdateProgress($"步骤3/3：正在合成第 {completed}/{totalPages} 页...");
                }
            });

            UpdateProgress("步骤3/3：正在合并所有页面...");
            MergePdfs(pageDir, "dual-", totalPages, outputPath, bookmarks);
        }

        private void MergePdfs(string pageDir, string prefix, int totalPages, string outputPath,
            ArrayList bookmarks = null)
        {
            var document = new iTextSharp.text.Document();
            var copy = new PdfCopy(document, new FileStream(outputPath, FileMode.Create));
            document.Open();

            try
            {
                for (int i = 1; i <= totalPages; i++)
                {
                    string pagePath = Path.Combine(pageDir, $"{prefix}{i}.pdf");
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

        #endregion
    }
}
