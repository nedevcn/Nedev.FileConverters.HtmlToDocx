# Nedev.HtmlToDocx

[![.NET 10](https://img.shields.io/badge/.NET-10.0-purple)](https://dotnet.microsoft.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![NuGet](https://img.shields.io/nuget/v/Nedev.HtmlToDocx.svg)](https://www.nuget.org/packages/Nedev.HtmlToDocx/)

High-performance HTML to DOCX converter built on .NET 10 with zero third-party dependencies.

## ✨ Features

- 🚀 **Extreme Performance** - Zero-allocation parsing using `Span<T>` and `Memory<T>`
- 📦 **Zero Dependencies** - Pure .NET implementation, no third-party libraries required
- 🎯 **AOT Support** - Native AOT compilation support for faster startup and smaller size
- 🔄 **Streaming** - Streaming conversion for large files with low memory footprint
- ⚡ **Parallel Processing** - Multi-core parallel acceleration for batch conversion
- 🎨 **CSS Support** - Conversion of common CSS styles
- 📊 **Complex Tables** - Support for tables, merged cells, and complex structures
- 🖼️ **Image Support** - Automatic download and embedding of web images

## 📦 Installation

### NuGet Package

```bash
dotnet add package Nedev.HtmlToDocx
```

### CLI Tool

```bash
dotnet tool install --global Nedev.HtmlToDocx.Cli
```

## 🚀 Quick Start

### Basic Usage

```csharp
using Nedev.HtmlToDocx;

// Simple conversion
string html = "<h1>Hello World</h1><p>This is a test.</p>";
byte[] docxBytes = html.ToDocx();
File.WriteAllBytes("output.docx", docxBytes);

// Using service
using var service = new HtmlToDocxService();
byte[] result = service.Convert(html);

// Async conversion
byte[] asyncResult = await html.ToDocxAsync();
```

### File Conversion

```csharp
using var service = new HtmlToDocxService();

// Synchronous
service.ConvertFile("input.html", "output.docx");

// Asynchronous
await service.ConvertFileAsync("input.html", "output.docx");
```

### CLI Usage

```bash
# Convert single file
htmltodocx input.html output.docx

# View help
htmltodocx --help
```

## ⚙️ Configuration Options

```csharp
var options = new ConverterOptions
{
    PreserveStyles = true,           // Preserve CSS styles
    DownloadImages = true,           // Download images
    MaxImageSize = 5 * 1024 * 1024,  // Maximum image size (5MB)
    ImageDownloadTimeout = TimeSpan.FromSeconds(30)  // Image download timeout
};

using var service = new HtmlToDocxService(options);
```

## 📋 Supported HTML Elements

### Text Elements
- `<p>` - Paragraph
- `<h1>` ~ `<h6>` - Headings
- `<br>` - Line break
- `<hr>` - Horizontal rule
- `<span>` / `<div>` - Containers

### Formatting
- `<strong>` / `<b>` - Bold
- `<em>` / `<i>` - Italic
- `<u>` - Underline
- `<s>` / `<strike>` - Strikethrough

### Links and Media
- `<a>` - Hyperlinks
- `<img>` - Images (supports data URI and HTTP URL)

### Lists
- `<ul>` / `<ol>` / `<li>` - Unordered/ordered lists

### Tables
- `<table>` / `<thead>` / `<tbody>` / `<tfoot>`
- `<tr>` / `<th>` / `<td>`
- `colspan` / `rowspan` - Cell merging

### Others
- `<blockquote>` - Blockquote
- `<pre>` / `<code>` - Code blocks

## 🎨 Supported CSS Properties

### Text Styles
- `color` - Font color
- `font-size` - Font size
- `font-family` - Font family
- `font-weight` - Font weight
- `font-style` - Font style
- `text-decoration` - Text decoration

### Paragraph Styles
- `text-align` - Text alignment
- `text-indent` - Text indent
- `line-height` - Line height
- `margin-top` / `margin-bottom` - Paragraph spacing

### Colors and Backgrounds
- `background-color` - Background color

### Box Model
- `width` - Width
- `padding` - Padding
- `border` - Border

## ⚡ Performance

Performance under typical workloads:

| File Size | Conversion Time | Throughput |
|-----------|-----------------|------------|
| 10 KB     | ~5 ms           | 2,000 chars/ms |
| 100 KB    | ~20 ms          | 5,000 chars/ms |
| 1 MB      | ~150 ms         | 6,600 chars/ms |

*Test environment: Intel i7-12700, 32GB RAM, .NET 10*

## 🏗️ Architecture

```
Nedev.HtmlToDocx/
├── Html/           # HTML Parsing
│   ├── HtmlParser      # Span<T>-based parser
│   ├── HtmlDocument    # DOM tree
│   └── HtmlEntities    # Entity decoding
├── Docx/           # DOCX Generation
│   ├── DocumentBuilder # Document builder
│   └── ZipArchiveHelper# ZIP compression
├── Conversion/     # Conversion Engine
│   └── HtmlToDocxConverter
└── HtmlToDocxService.cs  # Service interface
```

## 📝 TODO / Upcoming Features

### Medium Priority
- [ ] Page headers and footers
- [ ] Page numbering
- [ ] Table of contents generation
- [ ] Footnotes and endnotes
- [ ] Performance benchmarks suite

### Low Priority
- [ ] MathML support
- [ ] SVG to EMF conversion
- [ ] Batch processing GUI tool
- [ ] Changelog automation

## 🛠️ Development

### Build

```bash
cd src
dotnet build
```

### Test

```bash
dotnet test
```

### Publish AOT Version

```bash
cd src/Nedev.HtmlToDocx.Cli
dotnet publish -c Release -r win-x64 -p:PublishAot=true
```

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details

## 🤝 Contributing

Issues and Pull Requests are welcome!

## 📧 Contact

- GitHub: [https://github.com/nedev/HtmlToDocx](https://github.com/nedev/HtmlToDocx)
- NuGet: [https://www.nuget.org/packages/Nedev.HtmlToDocx](https://www.nuget.org/packages/Nedev.HtmlToDocx)
