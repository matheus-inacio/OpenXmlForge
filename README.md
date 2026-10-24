# OpenXmlForge

![CI](https://github.com/matheus-inacio/OpenXmlForge/workflows/CI/badge.svg)
[![NuGet](https://img.shields.io/nuget/dt/OpenXmlForge.svg)](https://www.nuget.org/packages/OpenXmlForge)
[![NuGet](https://img.shields.io/nuget/vpre/OpenXmlForge.svg)](https://www.nuget.org/packages/OpenXmlForge)

**OpenXmlForge** is a modern, high-performance .NET Word (`.docx`) template engine built on top of `DocumentFormat.OpenXml`.

Designed from the ground up for **modern .NET (.NET 10+)**, OpenXmlForge focuses on **low memory allocation, high throughput, and predictable resource usage** while keeping a clean, declarative `{{tag}}` template syntax.

## Installation

Install from [NuGet](https://www.nuget.org/packages/OpenXmlForge):

```bash
dotnet add package OpenXmlForge
```

## Quick Start

### Generate a `MemoryStream` (Async / Sync)

```csharp
using OpenXmlForge;

await using var templateStream = File.OpenRead("template.docx");

var data = new
{
    Order = new { Number = "ORD-2026-001", Date = DateTime.UtcNow },
    Customer = new { LegalName = "Acme Industrial Ltd", VIP = true },
    Items = new[]
    {
        new { Sku = "P-100", Description = "Industrial Servo Motor", Quantity = 2, Price = "$ 1,200.00" },
        new { Sku = "P-205", Description = "Frequency Inverter", Quantity = 1, Price = "$ 850.00" }
    }
};

// Returns a positioned MemoryStream (Position = 0) ready to be returned from an API or saved
await using MemoryStream output = await WordForge.GenerateStreamAsync(templateStream, data);
```

### Save Directly to a File or Stream

```csharp
using OpenXmlForge;

// Async file-to-file
await WordForge.SaveAsByTemplateAsync("output.docx", "template.docx", data);

// Sync file-to-file
WordForge.SaveAsByTemplate("output.docx", "template.docx", data);

// Extension method on any output Stream
await outputStream.SaveAsByTemplateAsync(templateStream, data);
```

---

## Template Features

### 1. Scalar & Nested Dotted Tags
OpenXmlForge accepts `IDictionary<string, object?>`, anonymous types, POCOs, and `ExpandoObject`. Both flat keys with dots (`["Customer.Name"]`) and nested object graphs (`new { Customer = new { Name = "..." } }`) resolve seamlessly:

```text
Invoice #{{Order.Number}}
Customer: {{Customer.LegalName}}
Notes: {{Document.Notes}}
```

Multiline strings (`\n`, `\r\n`) and `IEnumerable<string>` values automatically emit OpenXML `<w:br/>` line breaks while preserving run formatting.

---

### 2. Repeating Table Rows
Place collection tags (`{{Collection.Property}}`) inside a table row and OpenXmlForge will automatically clone the row for each item in `Collection`. Outer scalar variables (`{{Currency}}`, `{{Order.Number}}`) in the same row remain accessible:

| SKU | Description | Qty | Price |
| :--- | :--- | :--- | :--- |
| `{{Items.Sku}}` | `{{Items.Description}}` | `{{Items.Quantity}}` | `{{Currency}} {{Items.Price}}` |

### 3. Block & Inline Loops (`@foreach` / `{{foreach}}`)

**Block `@foreach`** (repeats paragraphs and tables between `@foreach` and `@endforeach`):
```text
@foreach(item in Items)
- {{item.Description}} ({{item.Quantity}}x)
@endforeach
```

**Inline `{{foreach}}`** (repeats text within a single paragraph):
```text
Tags: {{foreach item in Tags}}[{{item.Name}}] {{endforeach}}
```

Or using `WordForeach` with custom separators:
```csharp
using OpenXmlForge.Models;

var data = new Dictionary<string, object?>
{
    ["Tags"] = new WordForeach
    {
        Value = new[]
        {
            new { Tags = "Urgent" },
            new { Tags = "Priority" }
        },
        Separator = " | "
    }
};
```

### 4. Conditional Blocks & Inline Statements (`@if` / `{{if}}`)

Supports `==`, `!=`, `>`, `<`, `>=`, `<=` across numbers, `DateTime`, booleans, and strings.

**Block `@if`**:
```text
@if(Total > 1000)
Approved for Executive Discount
@endif
```

**Inline `{{if}}`**:
```text
Status: {{if(IsActive == true) ACTIVE endif}}{{if(IsActive == false) INACTIVE endif}}
```

### 5. Images (`WordPicture`)

Embed images from a `Stream` (including non-seekable streams), `byte[]`, or file path. When `Width` and `Height` are omitted, OpenXmlForge reads the natural dimensions from the image header and proportionally clamps the image to fit the containing table cell width if placed inside a table.

```csharp
using OpenXmlForge.Models;

var data = new Dictionary<string, object?>
{
    // Auto-sized from image header (and clamped to table cell width if inside a table)
    ["ProductImage"] = new WordPicture(imageStream),

    // Explicit dimensions in pixels
    ["CompanyLogo"] = new WordPicture("logo.png", width: 160, height: 60),

    // Floating / anchored image
    ["Watermark"] = new WordPicture(watermarkBytes, "png", width: 300, height: 300)
    {
        WrappingType = WordPictureWrappingType.Anchor,
        BehindDoc = true,
        HorizontalPositionOffset = 100,
        VerticalPositionOffset = 200
    }
};
```

### 6. Hyperlinks (`WordHyperLink`) & Colored Text (`WordColorText`)

```csharp
using OpenXmlForge.Models;

var data = new Dictionary<string, object?>
{
    ["PortalLink"] = new WordHyperLink
    {
        Url = "https://example.com/orders/1042",
        Text = "View Order in Portal",
        TargetFrame = TargetFrameType.Blank
    },
    ["StatusBadge"] = new WordColorText
    {
        Text = "APPROVED",
        FontColor = "#FFFFFF",
        HighlightColor = "#16A34A"
    }
};
```

## Building & Testing

```bash
dotnet build OpenXmlForge.slnx
dotnet test OpenXmlForge.slnx
```

## Why OpenXmlForge?

### Modern .NET Capabilities & Performance
- **Source-Generated Regex & Span Parsing**: Uses `[GeneratedRegex]`, `ReadOnlySpan<char>`, `ReadOnlySpan<byte>`, and `stackalloc` buffers for tag normalization and statement evaluation with minimal heap allocations.
- **Compiled Expression Tree Accessors**: POCO and anonymous object properties are accessed via cached, compiled `Expression.Lambda` delegates (`ConcurrentDictionary<Type, PropertyAccessor[]>`) instead of repeated runtime reflection, paired with `FrozenSet<T>` and `StringComparer.Ordinal` lookups.
- **Single-Pass DOM Tag Defragmenter**: Word, Google Docs, and LibreOffice frequently split `{{Variable.Name}}` tags across multiple `<w:r>` / `<w:t>` runs and inject `<w:proofErr>` spellcheck nodes. OpenXmlForge heals fragmented tags and `@foreach` / `@if` directives in a single pass directly over the OpenXML DOM—without serializing the document to strings or running full-document regex passes.

### Resource Efficiency & Streaming-First Design
- **Zero-Decode Image Header Inspection**: Extracts natural image dimensions (`PNG`, `JPEG`, `BMP`, `GIF`, `WebP`) and detects formats by magic bytes using `BinaryPrimitives` directly from the header—**zero pixel decoding** and no heavy imaging dependencies (`System.Drawing`, `SkiaSharp`, or `ImageSharp` are not required).
- **Non-Seekable & Network Stream Support**: Accepts non-seekable streams (such as `HttpClient` response streams or cloud object storage streams) for both templates and images using a lightweight `PrefixedStream` wrapper that replays the inspected header bytes before streaming the rest of the payload directly into the OpenXML package.
- **Concurrent Async Image Preparation & Deterministic Disposal**: Prepares image streams concurrently in async workflows and automatically disposes owned image streams when template generation finishes (configurable via `LeaveStreamOpen`).
- **Pre-Sized Output Buffers**: Estimates output `MemoryStream` capacity upfront when template length is known to avoid repeated large buffer reallocations.

### Advanced Layout & Document Correctness
- **Automatic Table Cell Image Clamping**: Measures usable cell width across explicit `dxa`, fractional Google Docs widths, `pct` percentage widths, `w:tblGrid` + `w:gridSpan`, and page margins—proportionally scaling oversized images so they never overflow table columns while keeping small icons at their natural size.
- **Unlimited Dotted Variables & Multi-Collection Rows**: Freely mix any number of dotted scalar paths (`{{Order.Number}}`, `{{Customer.LegalName}}`, `{{Address.City}}`) and repeating collection bindings in the same table row.
- **Vertical Merge (`<w:vMerge/>`) Row Grouping**: Automatically groups consecutive vertically merged rows and clones the entire group per collection item.
- **Isolated Shape & TextBox Scopes**: Processes `<w:txbxContent>` shapes from innermost to outermost so multiline text and tags inside `mc:AlternateContent` (`Choice` / `Fallback`) never leak into outer paragraphs.

## Acknowledgments

**OpenXmlForge** was inspired by [MiniWord](https://github.com/mini-software/MiniWord), reimagined from the ground up with a new architecture focused on modern .NET capabilities, streaming resource efficiency, and OpenXML layout accuracy.
