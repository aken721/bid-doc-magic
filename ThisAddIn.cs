using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace BidDocMagic
{
    public partial class ThisAddIn
    {
        public static Word.Application app;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            app = Globals.ThisAddIn.Application;
            PreloadNativeDlls();
        }

        private static readonly string NativeDir = Path.Combine(Path.GetTempPath(), "BidDocMagic");

        private static string GetDeploymentDir()
        {
            try
            {
                var codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase;
                var uri = new Uri(codeBase);
                return Path.GetDirectoryName(uri.LocalPath);
            }
            catch { }

            try
            {
                return Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            }
            catch { }

            return null;
        }

        private void PreloadNativeDlls()
        {
            string deployDir = GetDeploymentDir();
            if (string.IsNullOrEmpty(deployDir)) return;

            bool is64Bit = IntPtr.Size == 8;
            string srcName = is64Bit ? "pdfium_x64.dll" : "pdfium_x86.dll";
            string archSubDir = is64Bit ? "x64" : "x86";

            string[] candidates = new[]
            {
                Path.Combine(deployDir, srcName),
                Path.Combine(deployDir, archSubDir, "pdfium.dll"),
                Path.Combine(deployDir, "pdfium.dll"),
            };

            string srcPath = candidates.FirstOrDefault(p => File.Exists(p));
            if (srcPath == null)
            {
                System.Diagnostics.Debug.WriteLine($"BidDocMagic: pdfium.dll not found for {archSubDir}, searched: {string.Join("; ", candidates)}");
                return;
            }

            string nativeDir = NativeDir;
            try { Directory.CreateDirectory(nativeDir); } catch { }

            string dstPath = Path.Combine(nativeDir, "pdfium.dll");

            try
            {
                bool needCopy = !File.Exists(dstPath);
                if (!needCopy)
                {
                    var srcInfo = new FileInfo(srcPath);
                    var dstInfo = new FileInfo(dstPath);
                    needCopy = srcInfo.Length != dstInfo.Length;
                }
                if (needCopy)
                    File.Copy(srcPath, dstPath, true);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"BidDocMagic: Copy to {dstPath} failed: {ex.Message}");
            }

            string loadDir = File.Exists(dstPath) ? nativeDir : Path.GetDirectoryName(srcPath);

            try
            {
                SetDllDirectory(loadDir);
                string currentPath = Environment.GetEnvironmentVariable("PATH") ?? "";
                if (!currentPath.Contains(loadDir))
                    Environment.SetEnvironmentVariable("PATH", currentPath + ";" + loadDir);
            }
            catch { }

            try
            {
                string loadPath = File.Exists(dstPath) ? dstPath : srcPath;
                IntPtr hModule = LoadLibrary(loadPath);
                if (hModule == IntPtr.Zero)
                {
                    int err = Marshal.GetLastWin32Error();
                    System.Diagnostics.Debug.WriteLine($"BidDocMagic: LoadLibrary({loadPath}) failed, error={err}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"BidDocMagic: LoadLibrary({loadPath}) succeeded");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"BidDocMagic: LoadLibrary exception: {ex.Message}");
            }
        }


        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetDllDirectory(string lpPathName);

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
