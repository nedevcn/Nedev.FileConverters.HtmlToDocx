# Nedev.HtmlToDocx

[![.NET 10](https://img.shields.io/badge/.NET-10.0-purple)](https://dotnet.microsoft.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![NuGet](https://img.shields.io/nuget/v/Nedev.HtmlToDocx.svg)](https://www.nuget.org/packages/Nedev.HtmlToDocx/)

High-performance HTML to DOCX converter built on .NET 10 with zero third-party dependencies. Achieves high-fidelity conversion through a custom CSS engine and robust OpenXML generation.

## ✨ Features

- 🚀 **Extreme Performance** - Zero-allocation parsing using `Span<T>` and `ReadOnlySpan<char>`.
- 📦 **Zero Dependencies** - Pure .NET implementation, no third-party libraries (no OpenXML SDK).
- 🎨 **Full CSS Engine** - Custom cascading style resolution supporting `<style>` blocks and inline attributes.
- 📊 **Advanced Tables** - Sophisticated grid mapping for complex `colspan` and `rowspan` structures.
- 🔢 **Nested Lists** - Support for multi-level `<ul>` and `<ol>` with correct numbering indentation.
- 🖼️ **Multimedia** - Automatic embedding of local, remote (HTTP/S), and Base64 data URI images.
- 📐 **Page Layout** - Configurable page size (A4, Letter, etc.) and margins.
- 📄 **Headers & Footers** - HTML `<header>` and `<footer>` mapped to Word header/footer parts.
- 🔢 **Page Numbering** - Dynamic PAGE / NUMPAGES fields.
- 📖 **Table of Contents** - Auto-generated TOC field from headings.

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

// Using service with options
var options = new ConverterOptions { PageWidth = 11906, PageHeight = 16838 }; // A4
using var service = new HtmlToDocxService(options);
byte[] result = service.Convert(html);
```

### Async Conversion

```csharp
byte[] asyncResult = await html.ToDocxAsync();
```

## ⚙️ Configuration Options

```csharp
var options = new ConverterOptions
{
    PreserveStyles = true,           // Enable CSS engine
    DownloadImages = true,           // Download remote images
    MaxImageSize = 10 * 1024 * 1024, // 10MB
    PageWidth = 11906,               // A4 Width
    Margins = new PageMargins { Top = 1440, Bottom = 1440 }
};
```

## 📋 Supported HTML Elements

### Text & Formatting
- `<p>`, `<h1>`~`<h6>`, `<br>`, `<hr>`, `<span>`, `<div>`
- `<strong>`/`<b>`, `<em>`/`<i>`, `<u>`, `<s>`, `<font>`

### Links and Media
- `<a>` - Hyperlinks with relationship management.
- `<img>` - Supports `.jpg`, `.png`, `.gif` (local, remote, base64).

### Lists
- `<ul>` (Bullet) and `<ol>` (Numbered) with nested levels.

### Tables
- `<table>`, `<thead>`, `<tbody>`, `<tr>`, `<th>`, `<td>`.
- Full `colspan` / `rowspan` support with grid stabilization.

## 🎨 Supported CSS Properties

- **Typography**: `color`, `font-size` (pt/px), `font-family`, `font-weight`, `font-style`, `text-decoration`, `text-align`.
- **Layout**: `width`, `height`, `margin` (page-level).
- **Inheritance**: Styles correctly cascade from parent containers to text runs.

## ⚡ Performance

Verified performance on .NET 10:

| File Size | Conversion Time | Throughput |
|-----------|-----------------|------------|
| 10 KB     | ~3 ms           | 3,300 chars/ms |
| 100 KB    | ~15 ms          | 6,600 chars/ms |
| 1 MB      | ~120 ms         | 8,300 chars/ms |

*Test environment: Intel i7-12700, Windows 11, .NET 10.0*

## 🏗️ Architecture

```
Nedev.HtmlToDocx/
├── Html/           # Span-based HTML Tokenizer & Tree Builder
├── Css/            # Custom CSS Parser & Cascade Resolver
├── Docx/           # OpenXML (WordprocessingML) Generator
├── Models/         # Shared Converter & Layout Models
├── Utils/          # ZipArchiveHelper & IO Utilities
└── HtmlToDocxService.cs 
```

## 📝 Roadmap

### Medium Priority
- [ ] SVG to EMF conversion support

### Low Priority
- [ ] MathML support
- [ ] RTF export support

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details
