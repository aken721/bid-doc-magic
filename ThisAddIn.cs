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

        private void PreloadNativeDlls()
        {
            string asmDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            string[] nativeDlls = new string[] { "pdfium.dll" };

            string sourceDir = FindNativeDllDirectory();

            if (sourceDir != null && !string.IsNullOrEmpty(asmDir))
            {
                foreach (string dllName in nativeDlls)
                {
                    try
                    {
                        string srcPath = Path.Combine(sourceDir, dllName);
                        string dstPath = Path.Combine(asmDir, dllName);

                        if (File.Exists(srcPath))
                        {
                            bool needCopy = !File.Exists(dstPath);
                            if (!needCopy)
                            {
                                var srcInfo = new FileInfo(srcPath);
                                var dstInfo = new FileInfo(dstPath);
                                needCopy = srcInfo.Length != dstInfo.Length;
                            }

                            if (needCopy)
                            {
                                File.Copy(srcPath, dstPath, true);
                            }
                        }
                    }
                    catch { }
                }
            }

            if (!string.IsNullOrEmpty(asmDir))
            {
                SetDllDirectory(asmDir);
                string currentPath = Environment.GetEnvironmentVariable("PATH") ?? "";
                if (!currentPath.Contains(asmDir))
                {
                    Environment.SetEnvironmentVariable("PATH", currentPath + ";" + asmDir);
                }
            }

            foreach (string dllName in nativeDlls)
            {
                try
                {
                    string dllPath = Path.Combine(asmDir, dllName);
                    if (!File.Exists(dllPath) && sourceDir != null)
                    {
                        dllPath = Path.Combine(sourceDir, dllName);
                    }

                    if (File.Exists(dllPath))
                    {
                        IntPtr hModule = LoadLibrary(dllPath);
                        if (hModule == IntPtr.Zero)
                        {
                            int err = Marshal.GetLastWin32Error();
                            System.Diagnostics.Debug.WriteLine($"LoadLibrary({dllPath}) failed, error={err}");
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"Native DLL not found: {dllName}");
                    }
                }
                catch { }
            }
        }

        private string FindNativeDllDirectory()
        {
            var searchDirs = new List<string>();

            try
            {
                string asmDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                if (!string.IsNullOrEmpty(asmDir)) searchDirs.Add(asmDir);
            }
            catch { }

            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                if (!string.IsNullOrEmpty(baseDir)) searchDirs.Add(baseDir);
            }
            catch { }

            try
            {
                string asmDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                if (!string.IsNullOrEmpty(asmDir))
                {
                    string parentDir = Path.GetDirectoryName(Path.GetDirectoryName(asmDir));
                    if (!string.IsNullOrEmpty(parentDir))
                    {
                        searchDirs.Add(Path.Combine(parentDir, "Debug"));
                        searchDirs.Add(Path.Combine(parentDir, "Release"));
                        searchDirs.Add(Path.Combine(parentDir, "x64", "Debug"));
                        searchDirs.Add(Path.Combine(parentDir, "x64", "Release"));
                    }
                }
            }
            catch { }

            try
            {
                string asmDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                if (!string.IsNullOrEmpty(asmDir) && asmDir.Contains("AppData"))
                {
                    string[] knownPaths = new string[]
                    {
                        @"E:\SourceCode\Csharp\BidDocMagic-Free\BidDocMagic\bin\Debug",
                        @"E:\SourceCode\Csharp\BidDocMagic-Free\BidDocMagic\bin\x64\Debug",
                        @"E:\SourceCode\Csharp\BidDocMagic-Free\BidDocMagic\bin\Release",
                    };
                    foreach (string p in knownPaths)
                    {
                        if (Directory.Exists(p)) searchDirs.Add(p);
                    }
                }
            }
            catch { }

            foreach (string dir in searchDirs)
            {
                if (string.IsNullOrEmpty(dir)) continue;
                try
                {
                    if (File.Exists(Path.Combine(dir, "pdfium.dll")))
                        return dir;
                }
                catch { }
            }

            return null;
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
