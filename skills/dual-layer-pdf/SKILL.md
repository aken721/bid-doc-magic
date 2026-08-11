---
name: "dual-layer-pdf"
slug: "dual-layer-pdf"
version: "1.4.0"
displayName: "dual-layer-pdf"
summary: "一键将 Word / Excel / PowerPoint / PDF 转换为不可编辑、可搜索的双层 PDF——图片层覆盖在上，文字层隐藏于下，看起来像纯图片却保留全文检索能力；支持目录批量转换、PDFium 内置渲染，并可通过 --image-only 输出纯图片 PDF（无可搜索文本层），完美适配招投标、合同归档、电子凭证等"可打印但不可改"场景。"
description: "将 Word / Excel / PowerPoint / PDF 转换为不可编辑、可搜索的双层 PDF（图片覆盖层 + 文本层）的独立 CLI 工具。任何智能体安装后即可用一条命令完成转换。触发词：双层PDF, bid-pdf, dual-layer PDF, 不可编辑PDF, 纯图片PDF, docx/xlsx/pptx 转 PDF"
license: "MIT"
---

# BidDocMagic - 双层 PDF 转换器（通用技能）

一个**通用技能**，可将 Office 文档（.doc、.docx、.xls、.xlsx、.ppt、.pptx）转换为双层 PDF。图片层覆盖在上（OverContent），文本层隐藏于下——使 PDF 看起来像纯图片，同时保留全文可搜索性。

也支持通过 `--image-only` 参数输出**纯图片 PDF**（仅图片层、无文本层、不可搜索），适用于不需要可搜索性、仅需视觉呈现的场景。

任何智能体（work agent、code agent 等）安装本技能后，即可用一条命令完成文档转换。

## 适用场景

- 用户需要将 Word/Excel/PowerPoint 文档转换为不可编辑、可打印的双层 PDF
- 用户提到"双层PDF"、"bid-pdf"、"dual-layer PDF"、"不可编辑PDF"
- 用户需要 PDF 可打印但不可编辑
- 用户希望保护文档内容不被修改，同时保留文字可搜索性
- 用户需要输出纯图片 PDF（无可搜索文本层），提到"纯图片PDF"、"image-only"、"仅图片"

## 支持格式

| 应用 | 格式 | 转换方式 |
|------|------|----------|
| Word | .doc, .docx | `Word.Document.SaveAs2(wdFormatPDF)` |
| Excel | .xls, .xlsx | `Workbook.ExportAsFixedFormat(xlTypePDF)` |
| PowerPoint | .ppt, .pptx | `Presentation.SaveAs(ppSaveAsPDF)` |
| PDF | .pdf | 直接输入（使用 `--pdf-input` 跳过 Office 转换） |

## 快速开始

### 安装（下载预编译二进制）— 推荐

本技能在 SkillHub 上仅以文档形式发布。首次使用时，从 Gitee 下载预编译二进制文件（exe + 所有依赖）到技能的 `bin/` 目录。

> **已安装则跳过：** 如果 `bin/DualLayerPdfConverter.exe` 已存在，说明二进制就绪——跳过本步，直接进入[转换文档](#转换文档单命令)。

**安全说明：** 二进制文件锁定到固定的 git commit，下载后通过 SHA-256 校验。如果任何文件的哈希值不匹配，脚本将报错中止，不会使用该二进制。

| 文件 | SHA-256 |
|------|---------|
| DualLayerPdfConverter.exe | `c9cc1b91227fa0190d637b821a51e9e6ab81eba5908794fbc4168ebe82d1fcdc` |
| DualLayerPdfConverter.exe.config | `e7d2bd330d4602e7bbec3def01864e028bd41aa7270cd8f4147763559cf39073` |
| iTextSharp.dll | `0bdfc493f2975d8098615b8826f01494b14d2e2605b82d0a5580f93f19840693` |
| PDFiumSharp.dll | `2f9238237a2f150a3c6b0c7ee7281785aaca85efbbcfb5083faba5901b26d2e6` |
| PDFiumSharp.NativeBinaries.dll | `1e07b21d1c9ac7ab299b933a6b085180d1620b1eb4c6685f826bc6c31b105abf` |
| System.Buffers.dll | `c65fff603b283dc966d1a8b730c11d5e5e750e8021bd24640612f6cc3f2c6fb7` |
| System.Memory.dll | `bf3fb84664f4097f1a8a9bc71a51dcf8cf1a905d4080a4d290da1730866e856f` |
| System.Numerics.Vectors.dll | `1d3ef8698281e7cf7371d1554afef5872b39f96c26da772210a33da041ba1183` |
| System.Runtime.CompilerServices.Unsafe.dll | `bfc8f02a96934786ca5a0514a3b657021c12542e215e94b78fdcc74bfeffe3d3` |
| pdfium.dll | `4be3d3ad88a55a1e6545cf4c131d771eaf3fd3540142ff42955ebad7e68d989e` |
| pdfium_x64.dll | `4be3d3ad88a55a1e6545cf4c131d771eaf3fd3540142ff42955ebad7e68d989e` |
| pdfium_x86.dll | `0bd09148566cf558ef994ca0b3c1736769fd60bbbb2bf3bfc4ac47fcdd05fb23` |

源码 commit：`a33ab63df6f7d2862c461a2a9b9d4b1c1a186e5a`（已锁定，不可变）

```powershell
# 在技能目录下运行（即 SKILL.md 所在目录）
# 幂等 + 完整性校验：exe 已存在则跳过；下载后校验 SHA-256
$dest = ".\bin"
if (Test-Path "$dest\DualLayerPdfConverter.exe") {
  Write-Host "Binaries already installed at $dest\DualLayerPdfConverter.exe — skipping download."
} else {
  $commit = "a33ab63df6f7d2862c461a2a9b9d4b1c1a186e5a"
  $base = "https://gitee.com/aken721/bid-doc-magic/raw/$commit/skills/dual-layer-pdf/src/bin/Release"
  New-Item -ItemType Directory -Force -Path $dest | Out-Null
  $files = @{
    "DualLayerPdfConverter.exe"                      = "c9cc1b91227fa0190d637b821a51e9e6ab81eba5908794fbc4168ebe82d1fcdc"
    "DualLayerPdfConverter.exe.config"               = "e7d2bd330d4602e7bbec3def01864e028bd41aa7270cd8f4147763559cf39073"
    "iTextSharp.dll"                                 = "0bdfc493f2975d8098615b8826f01494b14d2e2605b82d0a5580f93f19840693"
    "PDFiumSharp.dll"                                = "2f9238237a2f150a3c6b0c7ee7281785aaca85efbbcfb5083faba5901b26d2e6"
    "PDFiumSharp.NativeBinaries.dll"                 = "1e07b21d1c9ac7ab299b933a6b085180d1620b1eb4c6685f826bc6c31b105abf"
    "System.Buffers.dll"                             = "c65fff603b283dc966d1a8b730c11d5e5e750e8021bd24640612f6cc3f2c6fb7"
    "System.Memory.dll"                              = "bf3fb84664f4097f1a8a9bc71a51dcf8cf1a905d4080a4d290da1730866e856f"
    "System.Numerics.Vectors.dll"                    = "1d3ef8698281e7cf7371d1554afef5872b39f96c26da772210a33da041ba1183"
    "System.Runtime.CompilerServices.Unsafe.dll"     = "bfc8f02a96934786ca5a0514a3b657021c12542e215e94b78fdcc74bfeffe3d3"
    "pdfium.dll"                                     = "4be3d3ad88a55a1e6545cf4c131d771eaf3fd3540142ff42955ebad7e68d989e"
    "pdfium_x64.dll"                                 = "4be3d3ad88a55a1e6545cf4c131d771eaf3fd3540142ff42955ebad7e68d989e"
    "pdfium_x86.dll"                                 = "0bd09148566cf558ef994ca0b3c1736769fd60bbbb2bf3bfc4ac47fcdd05fb23"
  }
  foreach ($f in $files.Keys) {
    Write-Host "Downloading $f ..."
    Invoke-WebRequest -Uri "$base/$f" -OutFile "$dest\$f" -UseBasicParsing
    $actual = (Get-FileHash -Path "$dest\$f" -Algorithm SHA256).Hash.ToLower()
    if ($actual -ne $files[$f]) {
      Remove-Item "$dest\$f" -Force
      Write-Error "SHA-256 mismatch for $f — download aborted. Expected $($files[$f]), got $actual"
      exit 1
    }
  }
  Write-Host "Done. All binaries verified. Binary: $dest\DualLayerPdfConverter.exe"
}
```

下载完成后，转换器位于 `bin/DualLayerPdfConverter.exe`，无需任何构建工具。

<details>
<summary>备选方案：从源码构建（需要 Visual Studio / MSBuild）</summary>

```powershell
cd {skill_dir}/src
nuget restore packages.config -PackagesDirectory packages
msbuild DualLayerPdfConverter.csproj /p:Configuration=Release /p:Platform=x64 /verbosity:minimal
# 输出：src/bin/Release/DualLayerPdfConverter.exe
```

如果 `nuget` 或 `msbuild` 不在 PATH 中：
- nuget：从 https://dist.nuget.org/win-x86-commandline/latest/nuget.exe 下载
- msbuild：通过 `vswhere -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe` 查找

</details>

### 转换文档（单命令）

> 下方命令中的 `{skill_dir}` 请替换为实际的技能目录路径（即 `SKILL.md` 所在目录）。

```bash
# 转换 Word 文档
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\path\to\document.docx"

# 转换 Excel 工作簿
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\path\to\spreadsheet.xlsx"

# 转换 PowerPoint 演示文稿
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\path\to\presentation.pptx"

# 带选项转换
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\path\to\document.docx" -o "C:\output\result.pdf" -d 300 -t 4

# 将已有 PDF 转换为双层
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\path\to\existing.pdf" --pdf-input

# 输出纯图片 PDF（无文本层，不可搜索）
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\path\to\document.docx" --image-only

# 将已有 PDF 转换为纯图片 PDF
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\path\to\existing.pdf" --pdf-input --image-only

# 批量转换目录下所有支持的 Office 文件
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\path\to\docs\"

# 批量转换目录下所有 PDF 文件
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\path\to\pdfs\" --pdf-input

# 批量转换为纯图片 PDF
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\path\to\docs\" --image-only

# 批量转换并指定输出目录
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\path\to\docs\" -o "C:\output\"
```

### 命令行选项

| 选项 | 缩写 | 默认值 | 说明 |
|------|------|--------|------|
| `--input` | `-i` | （必填） | 输入文件或目录路径 |
| `--output` | `-o` | `<输入>_DualPDF.pdf` 或 `<输入>_ImgPDF.pdf` | 输出：单文件时为路径，批量时为目录。若目标文件已存在，自动添加 `(01)`、`(02)` 等后缀 |
| `--dpi` | `-d` | `300` | 渲染 DPI（50-1200） |
| `--threads` | `-t` | CPU 核心数 | PDF 合成阶段最大并行线程数（渲染阶段为串行） |
| `--pdf-input` | | `false` | 将输入视为 PDF（跳过 Office 转 PDF 步骤） |
| `--image-only` | | `false` | 输出纯图片 PDF（仅图片层、无文本层、不可搜索）。默认输出文件名后缀为 `_ImgPDF.pdf` |
| `--open` | | `false` | 转换完成后打开 PDF（仅单文件） |

## 架构流程

```
输入 (.doc/.docx/.xls/.xlsx/.ppt/.pptx)
      │
      ▼
┌──────────────────┐
│ 步骤1：生成       │  Office 文档 → PDF（文本层）通过 Office Interop
│   文本层PDF       │  Word: SaveAs2 | Excel: ExportAsFixedFormat | PPT: SaveAs
│                   │  若使用 --pdf-input 则跳过
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ 步骤2：渲染       │  PDF → 高分辨率 PNG 图片（逐页）
│   页面图片        │  引擎：PDFium（唯一引擎，内置）
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│ 步骤3：合成       │  默认：文本 PDF + 图片覆盖层 → 双层 PDF（可搜索）
│   最终PDF         │  --image-only：仅图片 → 纯图片 PDF（不可搜索）
└──────────────────┘
```

### 输出模式对比

| 模式 | 参数 | 输出后缀 | 文本层 | 可搜索 | 适用场景 |
|------|------|----------|--------|--------|----------|
| 双层 PDF（默认） | （无） | `_DualPDF.pdf` | ✅ 隐藏在下层 | ✅ | 需要可搜索的不可编辑文档 |
| 纯图片 PDF | `--image-only` | `_ImgPDF.pdf` | ❌ 无 | ❌ | 仅需视觉呈现，无需搜索 |

## 依赖

### NuGet 包（通过 `nuget restore` 自动还原）

| 包名 | 版本 | 用途 |
|------|------|------|
| iTextSharp-LGPL | 4.1.6 | PDF 操作（stamper、覆盖层） |
| PDFiumSharp | 1.4660.0-alpha1 | PDF 渲染引擎（唯一引擎） |
| PDFiumSharp.NativeBinaries | 1.4660.0 | pdfium.dll 原生二进制 |
| System.Memory | 4.5.5 | PDFiumSharp 依赖 |
| System.Buffers | 4.5.1 | PDFiumSharp 依赖 |
| System.Runtime.CompilerServices.Unsafe | 4.5.3 | PDFiumSharp 依赖 |

### COM Interop 引用

| 引用 | 用途 |
|------|------|
| Microsoft.Office.Interop.Word | Word 文档 → PDF 转换 |
| Microsoft.Office.Interop.Excel | Excel 工作簿 → PDF 转换 |
| Microsoft.Office.Interop.PowerPoint | PowerPoint 演示文稿 → PDF 转换 |
| Microsoft.Office.Core | Office 共享类型（MsoTriState 等） |

### 原生 DLL

| DLL | 用途 | 部署方式 |
|-----|------|----------|
| pdfium.dll | PDFium 渲染引擎 | 来自 PDFiumSharp.NativeBinaries NuGet 包 |

### 系统要求

- Windows x64
- .NET Framework 4.8
- Microsoft Word（用于 .doc/.docx → PDF 转换）
- Microsoft Excel（用于 .xls/.xlsx → PDF 转换）
- Microsoft PowerPoint（用于 .ppt/.pptx → PDF 转换）
- 使用 `--pdf-input` 处理已有 PDF 时无需安装 Office

### 安全与完整性

预编译二进制托管在公开的 Gitee 仓库（`aken721/bid-doc-magic`）。为应对供应链风险，采取以下措施：

1. **锁定 commit**：下载链接引用固定的 git commit hash（`a33ab63...`），而非可变分支。该 commit 的内容不可变——仓库后续修改不会影响已发布的技能版本。

2. **SHA-256 校验**：每个下载文件都会计算哈希值，并与本 SKILL.md 中记录的预期值比对。上方的哈希表是发布到 SkillHub 的技能元数据的一部分，会在技能审核期间被检查。任何不匹配都会中止安装并删除该文件。

3. **可审计性**：锁定 commit 对应的完整源码可在 `https://gitee.com/aken721/bid-doc-magic/tree/a33ab63df6f7d2862c461a2a9b9d4b1c1a186e5a/skills/dual-layer-pdf/src` 查阅，供独立审查。

4. **信任假设**：Gitee 是第三方代码托管服务。需要更强保障的用户应审查锁定 commit 的源码并从源码构建（见上方备选方案）。

## 关键设计决策

1. **独立 CLI 工具**：与 VSTO/Word 插件解耦。作为控制台应用运行，任何智能体均可调用。

2. **多格式支持**：按扩展名自动识别文件类型，调用对应的 Office Interop API 导出 PDF。核心双层 PDF 合成逻辑（PDFiumSharp + iTextSharp）与格式无关。

3. **图片在上、文字在下**：使用 `GetOverContent()` 将渲染后的图片放置在文本层之上，使 PDF 看起来像纯图片，同时在下层保留文字可搜索性。

4. **PDFium 渲染引擎**：唯一渲染引擎。通过 NuGet 内置，无需外部依赖。渲染快速可靠。不存在其他引擎选项，也无"默认"概念——PDFium 就是引擎。

5. **临时目录**：所有中间文件存放于 `%TEMP%\BidDocMagic`，转换完成后自动清理。

6. **原生 DLL 预加载**：启动时自动发现并从已知位置加载原生 DLL。

## 项目结构

```
skills/dual-layer-pdf/
├── SKILL.md                          # 本文件
└── src/
    ├── DualLayerPdfConverter.csproj  # 项目文件
    ├── packages.config               # NuGet 包配置
    ├── Program.cs                    # CLI 入口
    └── DualLayerPdfEngine.cs         # 核心转换引擎（Word/Excel/PPT/PDF）
```

## 智能体集成指南

### 适用于 Code Agent / Work Agent

**步骤 1 — 首次安装：** 从 Gitee 下载预编译二进制（无需构建工具）。锁定到 commit `a33ab63`，SHA-256 校验。若 `bin/DualLayerPdfConverter.exe` 已存在则跳过：

```powershell
# 在技能目录下运行（即 SKILL.md 所在目录）。幂等 + 完整性校验。
$dest = ".\bin"
if (-not (Test-Path "$dest\DualLayerPdfConverter.exe")) {
  $base = "https://gitee.com/aken721/bid-doc-magic/raw/a33ab63df6f7d2862c461a2a9b9d4b1c1a186e5a/skills/dual-layer-pdf/src/bin/Release"
  New-Item -ItemType Directory -Force -Path $dest | Out-Null
  $files = @{
    "DualLayerPdfConverter.exe"="c9cc1b91227fa0190d637b821a51e9e6ab81eba5908794fbc4168ebe82d1fcdc"; "DualLayerPdfConverter.exe.config"="e7d2bd330d4602e7bbec3def01864e028bd41aa7270cd8f4147763559cf39073"; "iTextSharp.dll"="0bdfc493f2975d8098615b8826f01494b14d2e2605b82d0a5580f93f19840693"; "PDFiumSharp.dll"="2f9238237a2f150a3c6b0c7ee7281785aaca85efbbcfb5083faba5901b26d2e6"; "PDFiumSharp.NativeBinaries.dll"="1e07b21d1c9ac7ab299b933a6b085180d1620b1eb4c6685f826bc6c31b105abf"; "System.Buffers.dll"="c65fff603b283dc966d1a8b730c11d5e5e750e8021bd24640612f6cc3f2c6fb7"; "System.Memory.dll"="bf3fb84664f4097f1a8a9bc71a51dcf8cf1a905d4080a4d290da1730866e856f"; "System.Numerics.Vectors.dll"="1d3ef8698281e7cf7371d1554afef5872b39f96c26da772210a33da041ba1183"; "System.Runtime.CompilerServices.Unsafe.dll"="bfc8f02a96934786ca5a0514a3b657021c12542e215e94b78fdcc74bfeffe3d3"; "pdfium.dll"="4be3d3ad88a55a1e6545cf4c131d771eaf3fd3540142ff42955ebad7e68d989e"; "pdfium_x64.dll"="4be3d3ad88a55a1e6545cf4c131d771eaf3fd3540142ff42955ebad7e68d989e"; "pdfium_x86.dll"="0bd09148566cf558ef994ca0b3c1736769fd60bbbb2bf3bfc4ac47fcdd05fb23"
  }
  foreach ($f in $files.Keys) {
    Invoke-WebRequest -Uri "$base/$f" -OutFile "$dest\$f" -UseBasicParsing
    $actual = (Get-FileHash -Path "$dest\$f" -Algorithm SHA256).Hash.ToLower()
    if ($actual -ne $files[$f]) { Remove-Item "$dest\$f" -Force; Write-Error "SHA-256 mismatch: $f"; exit 1 }
  }
}
```

**步骤 2 — 转换文档：**

```bash
# 转换 Word
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\docs\report.docx" -d 300

# 转换 Excel
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\docs\spreadsheet.xlsx" -d 300

# 转换 PowerPoint
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\docs\presentation.pptx" -d 300

# 输出纯图片 PDF（无文本层，不可搜索）
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\docs\report.docx" --image-only

# 将已有 PDF 转换为纯图片 PDF
{skill_dir}/bin/DualLayerPdfConverter.exe -i "C:\docs\existing.pdf" --pdf-input --image-only
```

输出文件默认为 `<输入文件名>_DualPDF.pdf`（双层）或 `<输入文件名>_ImgPDF.pdf`（纯图片），位于输入文件所在目录。若目标文件已存在，会自动添加编号后缀（如 `filename_DualPDF(01).pdf`、`filename_DualPDF(02).pdf`）——永远不会覆盖已有文件。

### 错误处理

- 退出码 0：成功
- 退出码 1：参数错误
- 退出码 2：转换错误（文件未找到、Office 不可用、DLL 缺失等）

错误信息输出到 stderr。
