using System;
using System.Collections.Generic;

namespace Nedev.FileConverters.HtmlToDocx.Core.Models;

public class ConverterOptions
{
    public bool PreserveStyles { get; set; } = true;
    public bool DownloadImages { get; set; } = true;
    public int MaxImageSize { get; set; } = 5 * 1024 * 1024;
    public TimeSpan ImageDownloadTimeout { get; set; } = TimeSpan.FromSeconds(30);

    // Advanced Features
    public bool EnablePageNumbering { get; set; } = true;
    public int TOCLevels { get; set; } = 3;

    // Page Layout (values in twips, 1 inch = 1440)
    public int PageWidth { get; set; } = 11906; // A4
    public int PageHeight { get; set; } = 16838; // A4
    public PageMargins Margins { get; set; } = new();
}

public class PageMargins
{
    public int Top { get; set; } = 1440;
    public int Bottom { get; set; } = 1440;
    public int Left { get; set; } = 1440;
    public int Right { get; set; } = 1440;
}

public sealed class RunProperties
{
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Underline { get; set; }
    public string? Color { get; set; }
    public int FontSize { get; set; } // Points
    public string? FontFamily { get; set; }
}

public sealed class TableData
{
    public List<TableRow> Rows { get; } = new();
}

public sealed class TableRow
{
    public List<TableCell> Cells { get; } = new();
}

public sealed class TableCell
{
    public string Text { get; set; } = string.Empty;
    public int ColSpan { get; set; } = 1;
    public RowMergeType RowMerge { get; set; } = RowMergeType.None;
}

public enum RowMergeType
{
    None,
    Restart,
    Continue
}
