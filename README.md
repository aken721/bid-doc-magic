# BidDocMagic / 双层PDF转换器

[中文](#中文) | [English](#english)

---

<a id="english"></a>

## Overview

A dual-layer PDF generator that converts Word documents (.docx) or existing PDFs into non-editable, printable dual-layer PDFs. The image layer sits on top (OverContent), the text layer underneath — making the PDF appear as a pure image while preserving text searchability.

The project provides two usage modes:

- **Word Add-In (VSTO)**: Integrates into Microsoft Word Ribbon, one-click conversion from within Word
- **Skill / CLI**: A standalone skill that any agent (work agent, code agent, etc.) can install and use with a single command. Supports single-file and batch directory conversion.

## Features

- Convert DOCX to dual-layer PDF with one click / one command
- **Batch directory conversion**: Convert all Word/PDF files in a folder at once
- PDFium rendering engine (fast, built-in)
- Multi-threaded PDF composition for faster conversion
- Configurable DPI (150-1200, default 300)
- Support direct PDF input (skip Word-to-PDF step)
- Automatic native DLL discovery and preloading
- Clean temporary file management

## Architecture

```
Input (.docx/.pdf)
       │
       ▼
┌──────────────────┐
│ Step 1: Generate  │  DOCX → PDF (text layer) via Word Interop
│   Text Layer PDF  │  Skip if --pdf-input is set
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Step 2: Render    │  PDF → High-res PNG images (per page, parallel)
│   Page Images     │  Engine: PDFium (built-in, fast rendering)
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ Step 3: Compose   │  Text PDF + Image overlay → Final dual-layer PDF
│   Dual-Layer PDF  │  Image on top (OverContent), text underneath
└──────────────────┘
```

## Quick Start

### Word Add-In

1. Open `BidDocMagic.sln` in Visual Studio
2. Build the solution
3. Install the VSTO Add-In
4. Click the "Dual-Layer PDF" button in the Word Ribbon

### Skill / CLI

#### Install

**Option A: Install from pre-built package (Recommended)**

1. Download `dual-layer-pdf.zip` from the `skills/` directory
2. Extract to any location, e.g. `C:\Tools\dual-layer-pdf\`
3. The executable is at `bin/DualLayerPdfConverter.exe`

**Option B: Build from source**

```powershell
cd skills/dual-layer-pdf/src
nuget restore packages.config -PackagesDirectory packages
msbuild DualLayerPdfConverter.csproj /p:Configuration=Release /p:Platform=AnyCPU /verbosity:minimal
# Output: src/bin/Release/DualLayerPdfConverter.exe
```

#### Single File Conversion

```powershell
# Convert a DOCX file
DualLayerPdfConverter.exe -i "C:\docs\report.docx"

# Convert with options
DualLayerPdfConverter.exe -i "C:\docs\report.docx" -d 300 -t 4

# Convert an existing PDF to dual-layer
DualLayerPdfConverter.exe -i "C:\docs\existing.pdf" --pdf-input
```

#### Batch Directory Conversion

```powershell
# Convert all DOCX/DOC files in a folder
DualLayerPdfConverter.exe -i "C:\docs\"

# Convert all PDF files in a folder
DualLayerPdfConverter.exe -i "C:\pdfs\" --pdf-input

# Batch convert with custom output directory
DualLayerPdfConverter.exe -i "C:\docs\" -o "C:\output\"
```

When `-i` points to a directory, the tool automatically scans all `.docx`/`.doc` files (or `.pdf` files with `--pdf-input`) and converts them one by one. Files already named `*_DualPDF.pdf` are skipped to avoid duplicate conversion. If the output file already exists, it is automatically renamed with a numbered suffix, e.g. `filename_DualPDF(01).pdf`.

#### CLI Options

| Option | Alias | Default | Description |
|--------|-------|---------|-------------|
| `--input` | `-i` | (required) | Input file or directory path |
| `--output` | `-o` | `<input>_DualPDF.pdf` | Output: file path for single, directory for batch |
| `--dpi` | `-d` | `300` | Render DPI (150-1200) |
| `--threads` | `-t` | CPU core count | Max parallel threads per file |
| `--pdf-input` | | `false` | Treat input as PDF (skip Word-to-PDF step) |
| `--open` | | `false` | Open output PDF after conversion (single file only) |

## Project Structure

```
BidDocMagic/
├── Ribbon1.cs                    # Word Add-In: Ribbon UI + core conversion logic
├── PdfSettingsDialog.cs          # Word Add-In: Settings dialog + About
├── ProgressForm.cs               # Word Add-In: Progress bar with cancel support
├── ThisAddIn.cs                  # Word Add-In: VSTO entry point + DLL preloading
├── BidDocMagic.csproj            # Word Add-In project
├── skills/
│   ├── dual-layer-pdf.zip        # Pre-built skill package (ready to install)
│   └── dual-layer-pdf/
│       ├── SKILL.md              # Skill definition
│       └── src/                  # Source code
│           ├── Program.cs        # CLI entry point + batch mode
│           ├── DualLayerPdfEngine.cs  # Core conversion engine
│           ├── DualLayerPdfConverter.csproj
│           └── packages.config
├── LICENSE                       # License (Chinese)
├── LICENSE.en                    # License (English)
└── README.md                     # This file
```

## Dependencies

| Package | Version | License | Purpose |
|---------|---------|---------|---------|
| iTextSharp-LGPL | 4.1.6 | LGPL-2.1 | PDF manipulation |
| PDFiumSharp | 1.4660.0-alpha1 | Apache-2.0 | PDFium rendering |
| PDFiumSharp.NativeBinaries | 1.4660.0 | BSD-3-Clause | pdfium.dll native binary |
| System.Memory | 4.5.5 | MIT | Required by PDFiumSharp |
| System.Buffers | 4.5.1 | MIT | Required by PDFiumSharp |
| System.Runtime.CompilerServices.Unsafe | 4.5.3 | MIT | Required by PDFiumSharp |

## System Requirements

- Windows x86/x64
- .NET Framework 4.8
- Microsoft Word 2013+ (for DOCX → PDF conversion; not needed with `--pdf-input`)

## License

This project is licensed under **AGPL-3.0**.

See [LICENSE](LICENSE) (Chinese) / [LICENSE.en](LICENSE.en) (English) for details.

---

<a id="中文"></a>

## 概述

双层PDF生成器，可将 Word 文档（.docx）或已有 PDF 转换为不可编辑、可打印的双层PDF。图片覆盖层在上（OverContent），文本可搜索层在下——使 PDF 看起来像纯图片，同时保留文字可搜索性。

项目提供两种使用方式：

- **Word 插件（VSTO）**：集成到 Microsoft Word 功能区，在 Word 中一键转换
- **Skill / CLI 技能**：独立技能包，任何智能体（work agent、code agent 等）安装后即可使用一条命令转换。支持单文件和目录批量转换。

## 功能特性

- 一键/一条命令将 DOCX 转换为双层 PDF
- **目录批量转换**：一次性转换文件夹内所有 Word/PDF 文件
- PDFium 渲染引擎（内置，快速渲染）
- 多线程PDF合成，加速转换过程
- 可配置 DPI（150-1200，默认 300）
- 支持直接输入 PDF（跳过 Word 转 PDF 步骤）
- 自动发现和预加载原生 DLL
- 自动清理临时文件

## 架构流程

```
输入 (.docx/.pdf)
       │
       ▼
┌──────────────────┐
│ 步骤1：生成       │  DOCX → PDF（文本层）通过 Word Interop
│   文本层PDF       │  若使用 --pdf-input 则跳过
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ 步骤2：渲染       │  PDF → 高分辨率PNG图片（逐页，并行）
│   页面图片        │  引擎：PDFium（内置，快速渲染）
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ 步骤3：合成       │  文本PDF + 图片覆盖层 → 最终双层PDF
│   双层PDF         │  图片在上层（OverContent），文本在下层
└──────────────────┘
```

## 快速开始

### Word 插件

1. 在 Visual Studio 中打开 `BidDocMagic.sln`
2. 构建解决方案
3. 安装 VSTO 插件
4. 在 Word 功能区点击"双层PDF"按钮

### Skill / CLI 技能

#### 安装方式

**方式一：安装预构建包（推荐）**

1. 从 `skills/` 目录下载 `dual-layer-pdf.zip`
2. 解压到任意位置，如 `C:\Tools\dual-layer-pdf\`
3. 可执行文件位于 `bin/DualLayerPdfConverter.exe`

**方式二：从源码构建**

```powershell
cd skills/dual-layer-pdf/src
nuget restore packages.config -PackagesDirectory packages
msbuild DualLayerPdfConverter.csproj /p:Configuration=Release /p:Platform=AnyCPU /verbosity:minimal
# 输出：src/bin/Release/DualLayerPdfConverter.exe
```

#### 单文件转换

```powershell
# 转换单个 DOCX 文件
DualLayerPdfConverter.exe -i "C:\docs\report.docx"

# 带选项转换
DualLayerPdfConverter.exe -i "C:\docs\report.docx" -d 300 -t 4

# 将已有 PDF 转换为双层
DualLayerPdfConverter.exe -i "C:\docs\existing.pdf" --pdf-input
```

#### 目录批量转换

```powershell
# 转换文件夹内所有 DOCX/DOC 文件
DualLayerPdfConverter.exe -i "C:\docs\"

# 转换文件夹内所有 PDF 文件
DualLayerPdfConverter.exe -i "C:\pdfs\" --pdf-input

# 批量转换并指定输出目录
DualLayerPdfConverter.exe -i "C:\docs\" -o "C:\output\"
```

当 `-i` 指向目录时，工具自动扫描其中所有 `.docx`/`.doc` 文件（使用 `--pdf-input` 时扫描 `.pdf` 文件）并逐个转换。已命名为 `*_DualPDF.pdf` 的文件会自动跳过，避免重复转换。若输出文件已存在，会自动添加编号后缀，如 `文件名_DualPDF(01).pdf`。

#### 命令行选项

| 选项 | 缩写 | 默认值 | 说明 |
|------|------|--------|------|
| `--input` | `-i` | （必填） | 输入文件或目录路径 |
| `--output` | `-o` | `<输入>_DualPDF.pdf` | 输出：单文件时为路径，批量时为目录 |
| `--dpi` | `-d` | `300` | 渲染 DPI（150-1200） |
| `--threads` | `-t` | CPU核心数 | 每个文件的最大并行线程数 |
| `--pdf-input` | | `false` | 将输入视为PDF（跳过Word转PDF步骤） |
| `--open` | | `false` | 转换完成后打开PDF（仅单文件） |

## 项目结构

```
BidDocMagic/
├── Ribbon1.cs                    # Word插件：功能区UI + 核心转换逻辑
├── PdfSettingsDialog.cs          # Word插件：设置对话框 + 关于
├── ProgressForm.cs               # Word插件：进度条（支持取消）
├── ThisAddIn.cs                  # Word插件：VSTO入口 + DLL预加载
├── BidDocMagic.csproj            # Word插件项目
├── skills/
│   ├── dual-layer-pdf.zip        # 预构建技能包（可直接安装）
│   └── dual-layer-pdf/
│       ├── SKILL.md              # Skill 定义
│       └── src/                  # 源码
│           ├── Program.cs        # CLI 入口 + 批量模式
│           ├── DualLayerPdfEngine.cs  # 核心转换引擎
│           ├── DualLayerPdfConverter.csproj
│           └── packages.config
├── LICENSE                       # 开源协议（中文）
├── LICENSE.en                    # 开源协议（英文）
└── README.md                     # 本文件
```

## 依赖库

| 包名 | 版本 | 许可协议 | 用途 |
|------|------|----------|------|
| iTextSharp-LGPL | 4.1.6 | LGPL-2.1 | PDF操作 |
| PDFiumSharp | 1.4660.0-alpha1 | Apache-2.0 | PDFium渲染 |
| PDFiumSharp.NativeBinaries | 1.4660.0 | BSD-3-Clause | pdfium.dll 原生二进制 |
| System.Memory | 4.5.5 | MIT | PDFiumSharp依赖 |
| System.Buffers | 4.5.1 | MIT | PDFiumSharp依赖 |
| System.Runtime.CompilerServices.Unsafe | 4.5.3 | MIT | PDFiumSharp依赖 |

## 系统要求

- Windows x86/x64
- .NET Framework 4.8
- Microsoft Word 2013+（用于 DOCX → PDF 转换；使用 `--pdf-input` 时不需要）

## 开源协议

本项目基于 **AGPL-3.0** 开源。

详见 [LICENSE](LICENSE)（中文）/ [LICENSE.en](LICENSE.en)（英文）。
