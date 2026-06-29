---
name: "dual-layer-pdf"
description: "Convert DOCX to dual-layer PDF (image overlay + text layer) via a standalone CLI tool. Any agent can install and use it with a single command. Triggers: 双层PDF, bid-pdf, dual-layer PDF, 不可编辑PDF, convert docx to pdf"
---

# BidDocMagic - Dual-Layer PDF Converter (Universal Skill)

A **universal skill** that converts Word documents (.docx) to dual-layer PDFs. The image layer sits on top (OverContent), the text layer underneath — making the PDF appear as a pure image while preserving text searchability.

Any agent (work agent, code agent, etc.) can install this skill and convert documents with a single command.

## When to Use

- User wants to convert DOCX to non-editable, printable dual-layer PDFs
- User asks for "双层PDF", "bid-pdf", "dual-layer PDF", or "不可编辑PDF"
- User needs PDFs that are printable but not editable
- User wants to protect document content from modification while preserving text searchability

## Quick Start

### Install (One-time Setup)

```powershell
# 1. Restore NuGet packages
cd {skill_dir}/src
nuget restore packages.config -PackagesDirectory packages

# 2. Build with MSBuild
msbuild DualLayerPdfConverter.csproj /p:Configuration=Release /p:Platform=x64 /verbosity:minimal
```

This produces `DualLayerPdfConverter.exe` in `src/bin/Release/`.

> **Note**: If `nuget` or `msbuild` is not on PATH, use full paths:
> - nuget: download from https://dist.nuget.org/win-x86-commandline/latest/nuget.exe
> - msbuild: find via `vswhere -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe`

### Convert (Single Command)

```bash
# Convert a single DOCX file
DualLayerPdfConverter.exe -i "C:\path\to\document.docx"

# Convert with options
DualLayerPdfConverter.exe -i "C:\path\to\document.docx" -o "C:\output\result.pdf" -d 300 -t 4

# Convert an existing PDF to dual-layer
DualLayerPdfConverter.exe -i "C:\path\to\existing.pdf" --pdf-input

# Batch convert all DOCX files in a folder
DualLayerPdfConverter.exe -i "C:\path\to\docs\"

# Batch convert all PDF files in a folder
DualLayerPdfConverter.exe -i "C:\path\to\pdfs\" --pdf-input

# Batch convert with output directory
DualLayerPdfConverter.exe -i "C:\path\to\docs\" -o "C:\output\"
```

### Command-Line Options

| Option | Alias | Default | Description |
|--------|-------|---------|-------------|
| `--input` | `-i` | (required) | Input file or directory path |
| `--output` | `-o` | `<input>_DualPDF.pdf` | Output: file path for single, directory for batch. If the target file already exists, auto-renames with `(01)`, `(02)`, etc. suffix |
| `--dpi` | `-d` | `300` | Render DPI (50-1200) |
| `--threads` | `-t` | CPU core count | Max parallel threads for PDF composition only (rendering is serial) |
| `--pdf-input` | | `false` | Treat input as PDF (skip Word-to-PDF step) |
| `--open` | | `false` | Open output PDF after conversion (single file only) |

## Architecture Overview

```
Input (.docx)
      │
      ▼
┌──────────────────┐
│ Step 1: Generate  │  DOCX → PDF (text layer) via Microsoft.Office.Interop.Word
│   Text Layer PDF  │  Skip if --pdf-input is set
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Step 2: Render    │  PDF → High-res PNG images (per page)
│   Page Images     │  Engine: PDFium (the only engine, built-in)
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Step 3: Compose   │  Text PDF + Image overlay → Final dual-layer PDF
│   Dual-Layer PDF  │  Image on top (OverContent), text underneath
└──────────────────┘
```

## Dependencies

### NuGet Packages (auto-restored by `dotnet restore`)

| Package | Version | Purpose |
|---------|---------|---------|
| iTextSharp-LGPL | 4.1.6 | PDF manipulation (stamper, overlay) |
| PDFiumSharp | 1.4660.0-alpha1 | PDF rendering engine (the only engine) |
| PDFiumSharp.NativeBinaries | 1.4660.0 | pdfium.dll native binary |
| System.Memory | 4.5.5 | Required by PDFiumSharp |
| System.Buffers | 4.5.1 | Required by PDFiumSharp |
| System.Runtime.CompilerServices.Unsafe | 4.5.3 | Required by PDFiumSharp |

### Native DLLs

| DLL | Purpose | Deployment |
|-----|---------|------------|
| pdfium.dll | PDFium rendering engine | From PDFiumSharp.NativeBinaries NuGet package |

### System Requirements

- Windows x64
- .NET Framework 4.8
- Microsoft Word (for DOCX → PDF conversion; not needed if `--pdf-input`)

## Key Design Decisions

1. **Standalone CLI tool**: Decoupled from VSTO/Word AddIn. Runs as a console application that any agent can invoke.

2. **Image on top, text underneath**: Uses `GetOverContent()` to place the rendered image above the text layer, making the PDF appear as a pure image while preserving text searchability underneath.

3. **PDFium rendering engine**: The only rendering engine. Built-in via NuGet, no external dependencies required. Fast and reliable rendering. There is no other engine option and no "default" concept — PDFium is the engine.

4. **Temporary directory**: All intermediate files stored in `%TEMP%\BidDocMagic` and cleaned up after conversion.

5. **Native DLL preloading**: Automatically discovers and loads native DLLs from known locations at startup.

## Project Structure

```
skills/dual-layer-pdf/
├── SKILL.md                          # This file
└── src/
    ├── DualLayerPdfConverter.csproj  # Project file
    ├── packages.config               # NuGet packages
    ├── Program.cs                    # CLI entry point
    └── DualLayerPdfEngine.cs         # Core conversion engine
```

## Agent Integration Guide

### For Code Agent / Work Agent

After installing this skill, the agent can convert DOCX to dual-layer PDF by running:

```bash
# Build (first time only)
cd {skill_dir}/src
nuget restore packages.config -PackagesDirectory packages
msbuild DualLayerPdfConverter.csproj /p:Configuration=Release /p:Platform=x64 /verbosity:minimal

# Convert
{skill_dir}/src/bin/Release/DualLayerPdfConverter.exe -i "C:\docs\report.docx" -d 300
```

The output file defaults to `<input_filename>_DualPDF.pdf` in the same directory as the input. If the target file already exists, it is automatically renamed with a numbered suffix (e.g. `filename_DualPDF(01).pdf`, `filename_DualPDF(02).pdf`) — existing files are never overwritten.

### Error Handling

- Exit code 0: Success
- Exit code 1: Invalid arguments
- Exit code 2: Conversion error (file not found, Word not available, DLL missing, etc.)

Error messages are printed to stderr.
