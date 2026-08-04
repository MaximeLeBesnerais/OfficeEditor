# OfficeEditor

一套 .NET 9 库套件，用于**创建、编辑、生成和渲染 Office 文档** —— DOCX、PPTX、XLSX —— 无需 Office 或 LibreOffice。它提供流畅的 C# 构建器、JSON/YAML 指令集，以及一套经严格校验的声明式 JSON 词汇表，用于生成完整的 PPTX 演示文稿。渲染基于 **TypstBridge**，一个围绕 Typst 的原生 Rust 桥接：PPTX → PDF/PNG/SVG、DOCX → PDF、XLSX → PDF/PNG/SVG。

> **现已公开** —— OfficeEditor 0.7.1 以 MIT 许可证开源。

## 功能特性

- **三种格式，统一模型** —— Word (DOCX)、PowerPoint (PPTX)、Excel (XLSX)；支持从零创建，也支持编辑现有文件并保留样式
- **JSON 工作流** —— PPTX 拥有完整的声明式生成词汇表；**DOCX 支持声明式 JSON 生成**（流式 + 定位式两层、设计主题、语义化报告原型，参见 `docs/docx-generation.md`）；**XLSX 拥有丰富的指令/生成引擎**（类型化单元格、带填充/边框的命名样式、布局、表格），已接入 `officeeditor generate --output *.xlsx`，并可通过 Typst 管线渲染为 PDF/PNG/SVG（图片与行复制仍待实现）
- **渲染** —— 原生 TypstBridge（Typst 0.15.1）：PPTX PDF/PNG/SVG、DOCX PDF、XLSX PDF/PNG/SVG（公式仅渲染缓存的 `<v>` 值，不进行求值）；PPTX 相关接口会输出整份演示文稿的渲染计时
- **流畅的 C# API** —— `DocumentBuilder`、`PresentationBuilder`、`WorkbookBuilder`（支持文件、流或内存中的 `byte[]`）
- **指令集** —— JSON/YAML DOCX 操作、JSON PPTX 编辑操作，以及 v1 JSON XLSX 构建器词汇表
- **变量与邮件合并** —— 三种格式均支持 `{{variable}}` 检测与替换，另有 DOCX 批量合并
- **Markdown → DOCX** —— 通过 Markdig 进行富样式转换（1–6 级标题、嵌套强调、经 `http`/`https`/`mailto` 安全协议白名单并带内部锚点回退的真实超链接、图片、脚注、表格、任务列表、emoji、YAML front matter、自定义样式映射）
- **使用入口** —— 统一 CLI、ASP.NET Core API、MCP stdio 主机（4 个 `deck_*` 工具）以及 Web 演示应用
- **品牌档案** —— 从现有演示文稿中提取主题颜色/字体，形成可复用的令牌集

## 快速开始

### CLI

```bash
# Run from source (or install the tool — see Installation)
dotnet run --project OfficeEditor.Cli -- <command>

# Create documents (format auto-detected from extension)
officeeditor create output.docx --text "Hello World"
officeeditor create output.pptx --title "My Presentation"
officeeditor create output.xlsx --sheet "Sales"

# Generate a full deck from a JSON vocabulary
officeeditor generate demo/demo-deck.json --output deck.pptx

# Generate a DOCX report from a JSON vocabulary (design theme optional)
officeeditor generate report.json --output report.docx --theme corporate

# Generate a workbook from a JSON instruction set
officeeditor generate workbook.json --output workbook.xlsx

# Render an XLSX JSON instruction set to PDF (single file) or PNG (a directory of page-NNN.png pages)
officeeditor generate workbook.json --output workbook.pdf
officeeditor generate workbook.json --output workbook.png

# Detect variables in templates
officeeditor detect template.docx

# Merge template with data
officeeditor merge template.pptx data.json output.pptx

# DOCX-only Markdown → DOCX with a template, custom style map, and strict mode
dotnet run --project DocxEditor.Cli -- markdown guide.md guide.docx --template base.docx --style-map styles.json --strict

# DOCX-only instruction editing is available through the source CLI
dotnet run --project DocxEditor.Cli -- edit document.docx --instructions instructions.json
```

### C# API —— 从 JSON 生成演示文稿（旗舰路径）

```csharp
using PptxEditor.Core.Generation.Archetypes;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Schema;

var json = await File.ReadAllTextAsync("deck.json");   // "version": "2.0" vocabulary
var result = new GenerationDocumentParser().Validate(json);   // loud validator
if (!result.IsValid) { /* field-path errors with suggestions */ }

var doc        = ArchetypeExpander.Expand(result.Document!);
var components = ComponentExpander.Expand(doc);
var layout     = new LayoutResolver().Resolve(components);    // layout once…
var pptx       = new OoxmlEmitter().Emit(layout);             // …emit OOXML…
await File.WriteAllBytesAsync("deck.pptx", pptx.Bytes);
// …and the same layout feeds the Typst emitter for PDF/PNG/SVG previews.
```

### C# API —— 构建器（三种格式）

```csharp
using DocxEditor.Core.Builders;
using PptxEditor.Core.Builders;
using XlsxEditor.Core.Builders;

// DOCX
using var doc = DocumentBuilder.Create("report.docx");
doc.AddParagraph("Annual Report", "Heading1");
doc.AddMarkdown("This is **bold** and *italic*, with a list:\n- one\n- two");
doc.Save();

// PPTX — build, then render
using var deck = PresentationBuilder.Create("slides.pptx");
deck.AddSlide();
deck.CurrentSlide.AddTitle("Q4 Review").AddSubtitle("Sales");
deck.AddSlide();
deck.CurrentSlide.AddTitle("Numbers").AddTable(new List<List<string>>
{
    new() { "Product", "Q1", "Q2" },
    new() { "Widget", "100", "200" }
});
deck.Save();

byte[] pdf = deck.ExportToPdf();                                   // whole-deck PDF
byte[][] pngs = deck.ExportThumbnails(new ThumbnailOptions { Ppi = 150 });
string typ = deck.ExportToTypst();                                 // Typst source

// XLSX
using var book = WorkbookBuilder.Create("data.xlsx");
var sheet = book.AddWorksheet("Sales");
sheet.AddHeaderRow(new List<string> { "Product", "Q1", "Q2" })
     .AddDataRow(new List<string> { "Widget", "100", "200" }, 2)
     .AddFormulaRow(new List<string> { "Total", "=SUM(B2:B2)", "=SUM(C2:C2)" }, 3);
sheet.SetColumnWidth("A", 24).SetRowHeight(1, 28);
sheet.MergeCells("A4:C4");
book.Save();

// In-memory (services, Azure Functions, APIs)
using var mem = DocumentBuilder.Create();
mem.AddParagraph("Hello");
BinaryOfficeDocument bin = mem.ToBinaryDocument();   // bin.Bytes, bin.ContentType
```

### 变量与邮件合并

```csharp
var variables = builder.DetectVariables();          // find {{vars}} — all formats
builder.MergeVariables(new Dictionary<string, string>
{
    ["clientName"] = "Acme Corp",
    ["date"] = "2026-01-31"
});

// Lower-level template-engine classes also exist for conditional/loop expansion,
// but they are not wired into the unified CLI or public demo surfaces.
```

## 安装

NuGet 包当前版本为 **0.7.1**：

```bash
dotnet add package MaximeLB.PptxEditor.Core --version 0.7.1   # PPTX
dotnet add package MaximeLB.DocxEditor.Core --version 0.7.1   # DOCX
dotnet add package MaximeLB.XlsxEditor.Core --version 0.7.1   # XLSX
dotnet tool install -g MaximeLB.OfficeEditor.Cli --version 0.7.1   # `officeeditor` command
```

| 包 | 内容 |
|---|---|
| `MaximeLB.OfficeEditor.Core` | 共享服务（TypstCompilerService）、变量、模型 |
| `MaximeLB.DocxEditor.Core` | DOCX 构建器、指令、markdown、DOCX→Typst 转换器 |
| `MaximeLB.PptxEditor.Core` | PPTX 构建器、转换器、生成管线、品牌档案 |
| `MaximeLB.XlsxEditor.Core` | XLSX 构建器/读取 API、变量、JSON 指令 |
| `MaximeLB.TypstBridge.Managed` | 托管封装 + 原生 TypstBridge（osx-arm64、linux-x64、win-x64） |
| `MaximeLB.OfficeEditor.Cli` | 统一的 `officeeditor` CLI（dotnet 工具） |

或从源码构建：

```bash
git clone https://github.com/MaximeLeBesnerais/OfficeEditor.git
cd OfficeEditor && dotnet build        # 0 warnings, 0 errors (enforced)
```

## 仓库结构 —— 各部分所在位置

| 内容 | 位置 |
|---|---|
| **参考语料库**（许可证干净的夹具 + PowerPoint/Word 渲染的基准 PDF） | `examples/REF/` —— PPTX：`sales_acceleration_deck`（主参考，16 页，含 SmartArt）、`AetherLink-Glass-Shareholder-Overview`（带样式，品牌挖掘）、`northwind-investor-40`（压力测试）、`northwind-launch-review`（路演）、`northwind-demo`（自生成冒烟测试演示文稿）。DOCX：`annual-report`、`monitoring-report`（`{{variable}}` 模板） |
| **生成 JSON 示例** | `demo/demo-deck.json`（15 页 + `demo/themes.json` 预设）、`demo/deck.json`（6 页）、`decks/repo-overview.json` |
| **设计令牌集**（挖掘自品牌档案） | `PptxEditor.Core/Generation/Design/` |
| **一致性测试夹具**（逐原语生成测试） | `PptxEditor.Core/Generation/Fixtures/` |
| **工具** | `tools/convert-pptx`、`tools/convert-docx`、`tools/convert-xlsx`、`tools/visual-diff`、`tools/pptx-benchmark` |
| **代理/贡献者规则** | `AGENTS.md` |

## 入口

| 组件 | 运行方式 |
|---|---|
| 统一 CLI | `dotnet run --project OfficeEditor.Cli -- <command>` |
| API | `dotnet run --project OfficeEditor.Api --urls http://localhost:5001` |
| Web 演示 | `make dev` → http://localhost:5173/（`/demo` 即应用本身） |
| MCP 主机（JSON-RPC stdio） | `dotnet run --project OfficeEditor.Mcp` |
| 示例程序 | `dotnet run --project examples` |
| 转换 / 对比 / 基准测试工具 | `dotnet run --project tools/convert-pptx -- <in> <out> [--format pdf\|png\|svg\|typ]` · `dotnet run --project tools/convert-docx -- <in> <out> [--format pdf\|png\|svg\|typ]` · `dotnet run --project tools/convert-xlsx -- <in> <out> [--format pdf\|png\|svg\|typ\|json]` · `dotnet run --project tools/visual-diff -- --suite pptx\|gen\|xlsx` · `dotnet run --project tools/pptx-benchmark` |
| CLI 演示 | `make -f Makefile.demo demo`（预检 → 转换 REF 演示文稿 → 生成演示文稿，打印计时并打开 PDF） |

### 演示

`make dev`，打开 http://localhost:5173/ —— 四个标签页，所有计时均由服务端测量：

1. **渲染** —— 从白名单中选择 REF 演示文稿 → 计时 PNG/SVG 图库（当前基准：16 页销售演示文稿在 Apple Silicon 上约 329 ms 总体 / 每页 20.6 ms）
2. **生成** —— 实时编辑演示文稿标题 + 切换主题预设 → PPTX + 预览 + **可下载的 .pptx**
3. **任意渲染** —— 上传任意 `.pptx`
4. **对比** —— OfficeEditor 引擎 vs 无头 LibreOffice，幻灯片与计时并排对比（通常快 10× 以上）

API 与 Web 客户端是本地演示应用，并非生产级多租户服务。它们没有完整的身份验证、配额、沙箱或租户隔离层。参见 [SECURITY.md](SECURITY.md)。

## 架构

```
DocxEditor/                          # repo folder (historical name; product is OfficeEditor)
├── DocxEditor.Core/                 # DOCX: Builders, Content, Markdown, Instructions, Converters
├── PptxEditor.Core/                 # PPTX: Builders, Converters, Variables
│   └── Generation/                  #   JSON vocab: Schema (loud validator) → Archetypes → Components
│       #                            #   → Layout (once, pure C#) → Emit/OOXML + Emit/Typst (twice)
├── XlsxEditor.Core/                 # XLSX: Builders/read API, Variables, JSON Instructions
├── OfficeEditor.Core/               # Shared: TypstCompilerService, Variables, Exceptions
├── OfficeEditor.Cli/                # Unified multi-format CLI
├── OfficeEditor.Api/                # ASP.NET Core: deck sessions, previews, generation, demo/compare
├── OfficeEditor.Mcp/                # MCP stdio host: deck_anatomize, deck_replace_element, …
├── OfficeEditor.Web.Client/         # React/Vite/Tailwind demo app (/demo)
├── TypstBridge/                     # Rust native bridge + managed wrapper (Typst 0.15.1)
├── examples/                        # Sample programs + REF corpus
├── demo/, decks/                    # Generation JSON decks
└── tools/                           # convert-pptx, convert-docx, convert-xlsx, visual-diff, pptx-benchmark
```

关键设计决策：

- **布局一次，输出两次** —— 一次 C# 布局过程同时供 OOXML 发射器（交付）与 Typst 发射器（预览）使用；无需第二个布局引擎
- **每个原语都附带两个发射器 + 一个一致性测试夹具** —— 不存在测试不完整的功能
- **Typst 预览是模糊 OOXML 渲染的规范依据**（spec of record）
- **样式保留** —— 编辑时绝不修改现有样式定义；样式按 ID 引用

## API 与 MCP 接口

**API**（`OfficeEditor.Api`）：演示文稿上传/会话、逐页预览（`png|svg`，带 ETag 缓存）、演示文稿解剖（anatomy）、编辑指令、带计时的 JSON 生成（`generationMilliseconds`、`totalMilliseconds`）、演示端点（计时 REF 渲染、上传渲染、OfficeEditor 与 LibreOffice 对比），以及用于一次性转换的 `/api/convert` —— 通过声明式生成器将 JSON → DOCX/XLSX（空 JSON 生成空白文档）、Markdown → DOCX、DOCX → PDF、PPTX → PDF/PNG/SVG、XLSX → PDF/PNG/SVG（PNG/SVG 仅返回第一页）。

**MCP**（`OfficeEditor.Mcp`，stdio JSON-RPC）：`deck_anatomize`、`deck_replace_element`、`deck_render_slide`、`deck_generate`。

## 测试

```bash
dotnet test                 # full suite: 2,300+ cases across 5 test projects
```

- xUnit；生成管线采用逐原语一致性测试夹具，并以 RMSE 阈值校验
- 覆盖率：跨所有测试项目的合并并集行覆盖率（`scripts/check-coverage.py`）；CI 门槛下限 **83%** —— 权威来源是 `.github/workflows/ci.yml` 中的 `COVERAGE_THRESHOLD`
- 依赖 Typst 的测试通过环境变量开关：`OE_RUN_TYPST_COMPILE_TESTS=1 dotnet test`
- CI：在 PR 与 `main` 推送上运行 `build-test` —— Release 构建（警告视为错误）+ 完整测试套件 + 覆盖率门槛（83% 下限）；仅在 C# 相关路径变更时运行
- 视觉回归：PPTX 与生成测试套件目前已可运行。DOCX 套件已接入，但仍需在 `examples/output/ref/docx/` 下生成参考输出

## 性能

PPTX 渲染管线（PptxEditor → Typst → PNG/PDF）与无头 LibreOffice 对比
（`tools/pptx-benchmark`，Apple Silicon 上 5 次热运行的中位数 —— 按数量级看待）：

| 演示文稿 | 页数 | OfficeEditor 热运行（总计） | 每页 | LibreOffice 热运行（总计） | 每页 | 加速比 |
|---|---|---|---|---|---|---|
| sales_acceleration_deck | 16 | 329.4 ms | 20.6 ms | 5,924.0 ms | 370.2 ms | ~18.0× |
| AetherLink-Glass-Shareholder-Overview | 15 | 417.4 ms | 27.8 ms | 15,940.3 ms | 1,062.7 ms | ~38.2× |
| northwind-launch-review | 12 | 102.5 ms | 8.5 ms | 2,034.2 ms | 169.5 ms | ~19.8× |
| northwind-investor-40 | 40 | 322.2 ms | 8.1 ms | 5,629.7 ms | 140.7 ms | ~17.5× |

两边的产物一致：OfficeEditor 在一次编译中原生渲染幻灯片 PNG（一次打开即输出整份演示文稿的 PNG，@150ppi）；LibreOffice 无法直接栅格化 PPTX，因此其总耗时等于 `soffice --convert-to pdf` **再加上 `pdftoppm` 以 150dpi 栅格化** —— 而栅格化是主要开销。冷启动（全新进程，JIT + 后端探测）：每份演示文稿约 310–830 ms，视演示文稿而定。
复现方式：`dotnet run --project tools/pptx-benchmark`（方法论见 `tools/pptx-benchmark/README.md`）。

## Typst 编译后端

`TypstCompilerService` 通过 **TypstBridge**（进程内原生桥接，Typst 0.15.1）进行编译，并以外部 `typst` CLI 作为安全网。支持 PDF/SVG/PNG、多页输出、工作目录资源、显式字体路径（`--font-path`）、PNG PPI、持久化编译会话以及诊断信息。

## 状态与限制

**版本：** 0.7.1（pre-1.0）。公共 API 可能变动；在 v1.0 之前，TypstBridge ABI 有意保持流动（fluid）。

**平台：** 纯托管 .NET 9 + 按 RID 分发的原生 TypstBridge（osx-arm64、linux-x64、win-x64 在 CI 中构建）。
- **macOS (arm64)** —— 开发平台；所有功能均在此验证
- **linux-x64** —— 已测试
- **Windows** —— 预期可用；CI 中的运行时验证仍待完成

**渲染（PPTX → Typst → PDF/PNG/SVG）：**
- 文本、图片、形状与表格的渲染保真度良好
- 簇状条形图/柱形图可渲染；其他图表类型可能回退为占位符。SmartArt 由预渲染的绘图形状渲染，存在已文档化的主题色限制。动画不在静态预览模型范围内
- 字体保真度因平台而异；安装 Microsoft Office 字体（Aptos、Calibri）可提高准确度，但无法保证与 PowerPoint 完全一致（字体度量、断行、布局引擎均不相同）

**NuGet：** 包当前版本为 0.7.1。

## 安全

OfficeEditor 库在调用方进程中运行，并非用于防御恶意文档的沙箱。任何接受不受信任上传的部署都必须自行提供隔离、资源限制、身份验证和文件系统策略。API/Web 演示以及 LibreOffice 对比路径仅为本地便利功能，不属于产品安全边界。报告方式与完整信任模型参见 [SECURITY.md](SECURITY.md)。

## 许可证

MIT — Copyright 2026 Maxime Le Besnerais

本产品附带第三方组件，这些组件遵循各自的许可证；参见 [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md)（尤其是 Typst，Apache-2.0）。

## 贡献

本仓库采用受保护的 `main` / 工作分支 `dev` 的分支模型：

1. 从 `dev` 创建分支；CI 会在你的 PR 上运行（构建 + 完整测试 + 覆盖率门槛；仅在 C# 相关路径变更时触发）
2. 进入 `main` 的 PR 必须来自 `dev`（由 `guard-main` 强制执行）
3. 发布策略：从经过测试的 `main` 提交创建 `v*` 标签 → 原生构建矩阵 → NuGet 受信发布

工程规范与发布流程参见 `AGENTS.md`。

## 作者

**Maxime Le Besnerais**
