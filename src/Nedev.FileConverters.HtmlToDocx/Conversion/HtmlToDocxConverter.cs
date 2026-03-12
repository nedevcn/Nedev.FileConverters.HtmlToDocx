using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Text;
using Nedev.FileConverters.HtmlToDocx.Core.Docx;
using Nedev.FileConverters.HtmlToDocx.Core.Html;
using Nedev.FileConverters.HtmlToDocx.Core.Css;
using Nedev.FileConverters.HtmlToDocx.Core.Models;

namespace Nedev.FileConverters.HtmlToDocx.Core.Conversion;

public class HtmlToDocxConverter : IDisposable
{
    private readonly ConverterOptions _options;
    private readonly Stack<(int NumId, int Level)> _listStack = new();
    private readonly HttpClient _httpClient = new();
    private const int MaxNumberingLevel = 8;
    private const int ListBaseIndentLeft = 720;
    private const int ListHangingIndent = 360;
    private static readonly string[] BulletLvlText =
    {
        "", // Symbol bullet
        "o", // ASCII 'o'
        "", // Wingdings square bullet
        "·",
        "o",
        "",
        "·",
        "o",
        ""
    };

    private static readonly string[] BulletFonts =
    {
        "Symbol",
        "Courier New",
        "Wingdings",
        "Symbol",
        "Courier New",
        "Wingdings",
        "Symbol",
        "Courier New",
        "Wingdings"
    };

    private static string GetBulletChar(int level) => BulletLvlText[Math.Min(level, BulletLvlText.Length - 1)];
    private static string GetBulletFont(int level) => BulletFonts[Math.Min(level, BulletFonts.Length - 1)];
    private static int GetListLeftIndent(int level) => ListBaseIndentLeft * (level + 1);
    private static int GetListTabPos(int level) => GetListLeftIndent(level);

    public HtmlToDocxConverter(ConverterOptions? options = null)
    {
        _options = options ?? new ConverterOptions();
    }

    // convert css length values to twips (1pt = 20 twips)
    private static int ParseLengthToTwips(string length, int basePt = 12)
    {
        // reuse existing font-size parser with px->pt conversion then multiply by 20
        var pts = ParseFontSize(length, basePt);
        return pts * 20;
    }

    public byte[] Convert(string html)
    {
        var document = HtmlDocument.Parse(html);
        if (_options.PreserveStyles)
        {
            StyleResolver.ResolveStyles(document);
        }

        var builder = new DocumentBuilder();
        ConvertNode(document.Root, builder);
        return BuildDocx(builder.Build(_options), 
                        builder.BuildRelationships(), 
                        builder.BuildMedia(),
                        builder.BuildStyles(),
                        builder.BuildFonts(),
                        builder.BuildHeader(),
                        builder.BuildFooter());
    }

    public async Task<byte[]> ConvertAsync(string html, CancellationToken cancellationToken = default)
    {
        return await Task.Run(() => Convert(html), cancellationToken);
    }

    private void ConvertNode(HtmlNode node, DocumentBuilder builder)
    {
        switch (node.TagName.ToLowerInvariant())
        {
            case "h1":
            case "h2":
            case "h3":
            case "h4":
            case "h5":
            case "h6":
                var level = int.TryParse(node.TagName.Substring(1), out var l) ? l : 1;
                var (beforeH, afterH) = GetMarginSpacing(node);
                var (padLeftH, padRightH) = GetPaddingSpacing(node);
                var bgH = node.ComputedStyle != null && node.ComputedStyle.TryGetValue("background-color", out var bgc) ? ParseColor(bgc) : null;
                builder.StartParagraph($"Heading{level}", GetTextAlign(node), spacingBeforeTwips: beforeH, spacingAfterTwips: afterH, indentLeft: padLeftH, indentRight: padRightH, shdColor: bgH);
                foreach (var child in node.Children) ConvertNode(child, builder);
                builder.EndParagraph();
                return;
            case "p":
                var (beforeP, afterP) = GetMarginSpacing(node);
                var (padLeftP, padRightP) = GetPaddingSpacing(node);
                var bgP = node.ComputedStyle != null && node.ComputedStyle.TryGetValue("background-color", out var bgc2) ? ParseColor(bgc2) : null;
                builder.StartParagraph(textAlign: GetTextAlign(node), spacingBeforeTwips: beforeP, spacingAfterTwips: afterP, indentLeft: padLeftP, indentRight: padRightP, shdColor: bgP);
                foreach (var child in node.Children) ConvertNode(child, builder);
                builder.EndParagraph();
                return;
            case "table":
                ConvertTable(node, builder);
                return;
            case "ul":
                _listStack.Push((1, Math.Min(_listStack.Count, MaxNumberingLevel))); // 1 for Bullet
                foreach (var child in node.Children) ConvertNode(child, builder);
                _listStack.Pop();
                return;
            case "ol":
                _listStack.Push((2, Math.Min(_listStack.Count, MaxNumberingLevel))); // 2 for Numbered
                foreach (var child in node.Children) ConvertNode(child, builder);
                _listStack.Pop();
                return;
            case "li":
                if (_listStack.TryPeek(out var listInfo))
                {
                    var numId = listInfo.NumId;
                    var listLevel = Math.Min(listInfo.Level, MaxNumberingLevel);
                    int continuationIndentLeft = GetListLeftIndent(listLevel);

                    bool paragraphStarted = false;
                    bool markerUsed = false;

                    void StartListPara(HtmlNode? sourceForStyle = null)
                    {
                        if (paragraphStarted) return;

                        if (!markerUsed)
                        {
                            StartParagraphLikeNode(sourceForStyle, builder, listNumId: numId, listLevel: listLevel);
                            markerUsed = true;
                        }
                        else
                        {
                            StartParagraphLikeNode(
                                sourceForStyle,
                                builder,
                                indentLeft: continuationIndentLeft,
                                indentHanging: ListHangingIndent,
                                tabPos: continuationIndentLeft
                            );
                        }

                        paragraphStarted = true;
                    }

                    void EndParaIfAny()
                    {
                        if (!paragraphStarted) return;
                        builder.EndParagraph();
                        paragraphStarted = false;
                    }

                    void EnsureMarkerIfNeeded()
                    {
                        if (markerUsed) return;
                        StartParagraphLikeNode(null, builder, listNumId: numId, listLevel: listLevel);
                        builder.EndParagraph();
                        markerUsed = true;
                    }

                    foreach (var child in node.Children)
                    {
                        // Handle common block tags inside <li> as part of the same list item.
                        if (IsParagraphLikeBlock(child))
                        {
                            EndParaIfAny();
                            StartListPara(child);
                            foreach (var gc in child.Children) ConvertNode(gc, builder);
                            EndParaIfAny();
                            continue;
                        }

                        if (IsContainerBlock(child))
                        {
                            EndParaIfAny();
                            foreach (var gc in child.Children)
                            {
                                // treat container children as if they were direct children of <li>
                                if (IsParagraphLikeBlock(gc))
                                {
                                    StartListPara(gc);
                                    foreach (var gg in gc.Children) ConvertNode(gg, builder);
                                    EndParaIfAny();
                                }
                                else if (IsHardBlock(gc))
                                {
                                    EnsureMarkerIfNeeded();
                                    ConvertNode(gc, builder);
                                }
                                else
                                {
                                    StartListPara();
                                    ConvertNode(gc, builder);
                                }
                            }
                            continue;
                        }

                        if (IsHardBlock(child))
                        {
                            EndParaIfAny();
                            EnsureMarkerIfNeeded();
                            ConvertNode(child, builder);
                            continue;
                        }

                        // Inline content
                        StartListPara();
                        ConvertNode(child, builder);
                    }

                    EndParaIfAny();
                }
                else
                {
                    // Fallback if li is outside ul/ol
                    builder.StartParagraph();
                    foreach (var child in node.Children) ConvertNode(child, builder);
                    builder.EndParagraph();
                }
                return;
            case "br":
                builder.AddBreak();
                return;
            case "#text":
                builder.AddRun(node.Text, GetRunProperties(node));
                return;
            case "header":
                builder.SwitchToHeader();
                foreach (var child in node.Children) ConvertNode(child, builder);
                builder.SwitchToBody();
                return;
            case "footer":
                builder.SwitchToFooter();
                foreach (var child in node.Children) ConvertNode(child, builder);
                builder.SwitchToBody();
                return;
            case "div":
            case "span":
                var className = node.GetAttribute("class");
                if (className != null && className.Contains("page-number"))
                {
                    builder.AddPageNumberField();
                    return;
                }
                if (className != null && className.Contains("total-pages"))
                {
                    builder.AddTotalPagesField();
                    return;
                }
                if (node.TagName == "div" && (node.Id == "toc" || (className != null && className.Contains("toc"))))
                {
                    builder.AddTableOfContents(_options.TOCLevels);
                    return;
                }
                foreach (var child in node.Children) ConvertNode(child, builder);
                return;
            case "section":
            case "article":
            case "body":
            case "html":
            case "#document":
                // Container tags, just process children
                foreach (var child in node.Children) ConvertNode(child, builder);
                return;
            case "b":
            case "strong":
            case "i":
            case "em":
            case "u":
            case "font":
                // Inline tags, process children (runs will inherit styles via StyleResolver)
                foreach (var child in node.Children) ConvertNode(child, builder);
                return;
            case "a":
                var href = node.GetAttribute("href") ?? "#";
                builder.AddHyperlink(ExtractText(node), href);
                return;
            case "img":
                var src = node.GetAttribute("src");
                if (!string.IsNullOrEmpty(src))
                {
                    var imageData = ProcessImage(src).GetAwaiter().GetResult();
                    if (imageData != null)
                    {
                        builder.AddImage(imageData.Data, imageData.ContentType, imageData.Width, imageData.Height);
                    }
                }
                return;
            case "style":
            case "script":
            case "head":
                // Ignore these in the final document body
                return;
            default:
                // Fallback for unknown tags
                foreach (var child in node.Children) ConvertNode(child, builder);
                return;
        }
    }

    private void ConvertTable(HtmlNode node, DocumentBuilder builder)
    {
        var trNodes = node.Children.Where(c => c.TagName == "tr").ToList();
        if (!trNodes.Any()) trNodes = node.Children.SelectMany(c => c.Children).Where(c => c.TagName == "tr").ToList();
        if (!trNodes.Any()) return;

        int rowCount = trNodes.Count;
        // Accurate column count estimation is hard, use a safe buffer
        int initialColEstimate = trNodes.Max(tr => tr.Children.Count(c => c.TagName == "td" || c.TagName == "th")) * 2;
        var grid = new TableCell?[rowCount, initialColEstimate + 10];
        int maxCols = 0;

        for (int r = 0; r < rowCount; r++)
        {
            var tr = trNodes[r];
            var htmlCells = tr.Children.Where(c => c.TagName == "td" || c.TagName == "th").ToList();
            int cPos = 0;

            foreach (var cellNode in htmlCells)
            {
                // Find next free spot in the grid for this row
                while (cPos < grid.GetLength(1) && grid[r, cPos] != null) cPos++;
                if (cPos >= grid.GetLength(1)) break;

                int colspan = int.TryParse(cellNode.GetAttribute("colspan"), out var cs) ? cs : 1;
                int rowspan = int.TryParse(cellNode.GetAttribute("rowspan"), out var rs) ? rs : 1;

                var cell = new TableCell { 
                    Text = ExtractText(cellNode), 
                    ColSpan = colspan,
                    RowMerge = rowspan > 1 ? RowMergeType.Restart : RowMergeType.None
                };

                // cell styling
                if (cellNode.ComputedStyle != null)
                {
                    if (cellNode.ComputedStyle.TryGetValue("background-color", out var bgc))
                        cell.BackgroundColor = ParseColor(bgc);
                    var (pl, pr) = GetPaddingSpacing(cellNode);
                    cell.PaddingLeft = pl;
                    cell.PaddingRight = pr;
                    // top/bottom padding not used currently
                    if (cellNode.ComputedStyle.TryGetValue("padding-top", out var pt))
                        cell.PaddingTop = ParseLengthToTwips(pt);
                    if (cellNode.ComputedStyle.TryGetValue("padding-bottom", out var pb))
                        cell.PaddingBottom = ParseLengthToTwips(pb);
                }

                grid[r, cPos] = cell;
                maxCols = Math.Max(maxCols, cPos + colspan);

                // Mark grid occupancy for rowspan and colspan
                for (int i = 0; i < rowspan; i++)
                {
                    for (int j = 0; j < colspan; j++)
                    {
                        if (r + i < rowCount && cPos + j < grid.GetLength(1))
                        {
                            if (i == 0 && j == 0) continue; // Already set the primary cell
                            grid[r + i, cPos + j] = new TableCell { RowMerge = i > 0 ? RowMergeType.Continue : RowMergeType.None };
                        }
                    }
                }
                cPos += colspan;
            }
        }

        var tableData = new TableData();
        for (int r = 0; r < rowCount; r++)
        {
            var row = new TableRow();
            for (int c = 0; c < maxCols; c++)
            {
                var cell = grid[r, c];
                if (cell != null)
                {
                    row.Cells.Add(cell);
                    if (cell.ColSpan > 1) c += (cell.ColSpan - 1);
                }
                else
                {
                    // Fill gaps with empty cells to keep table consistent
                    row.Cells.Add(new TableCell());
                }
            }
            tableData.Rows.Add(row);
        }
        builder.AddTable(tableData);
    }

    private static string ExtractText(HtmlNode node)
    {
        if (node.TagName == "#text")
            return node.Text;
        
        var sb = new System.Text.StringBuilder();
        foreach (var child in node.Children)
        {
            sb.Append(ExtractText(child));
        }
        return sb.ToString();
    }

    private static (int before, int after) GetMarginSpacing(HtmlNode node)
    {
        int before = 0, after = 0;
        if (node.ComputedStyle != null)
        {
            if (node.ComputedStyle.TryGetValue("margin-top", out var mt))
                before = ParseLengthToTwips(mt);
            if (node.ComputedStyle.TryGetValue("margin-bottom", out var mb))
                after = ParseLengthToTwips(mb);
            if ((before == 0 && after == 0) && node.ComputedStyle.TryGetValue("margin", out var m))
            {
                var parts = m.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 1)
                {
                    before = after = ParseLengthToTwips(parts[0]);
                }
                else if (parts.Length == 2)
                {
                    before = ParseLengthToTwips(parts[0]);
                    after = ParseLengthToTwips(parts[1]);
                }
                else if (parts.Length == 3)
                {
                    before = ParseLengthToTwips(parts[0]);
                    after = ParseLengthToTwips(parts[2]);
                }
                else if (parts.Length == 4)
                {
                    before = ParseLengthToTwips(parts[0]);
                    after = ParseLengthToTwips(parts[2]);
                }
            }
        }
        return (before, after);
    }

    private static (int left, int right) GetPaddingSpacing(HtmlNode node)
    {
        int left = 0, right = 0;
        if (node.ComputedStyle != null)
        {
            if (node.ComputedStyle.TryGetValue("padding-left", out var pl))
                left = ParseLengthToTwips(pl);
            if (node.ComputedStyle.TryGetValue("padding-right", out var pr))
                right = ParseLengthToTwips(pr);
            if ((left == 0 && right == 0) && node.ComputedStyle.TryGetValue("padding", out var p))
            {
                var parts = p.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 1)
                {
                    left = right = ParseLengthToTwips(parts[0]);
                }
                else if (parts.Length == 2)
                {
                    left = right = ParseLengthToTwips(parts[1]);
                }
                else if (parts.Length == 3)
                {
                    left = ParseLengthToTwips(parts[1]);
                    right = ParseLengthToTwips(parts[1]);
                }
                else if (parts.Length == 4)
                {
                    left = ParseLengthToTwips(parts[3]);
                    right = ParseLengthToTwips(parts[1]);
                }
            }
        }
        return (left, right);
    }

    private static RunProperties GetRunProperties(HtmlNode node)
    {
        var props = new RunProperties();
        var style = node.ComputedStyle;

        if (style.TryGetValue("font-weight", out var weight))
            props.Bold = weight.Equals("bold", StringComparison.OrdinalIgnoreCase) || 
                         (int.TryParse(weight, out var w) && w >= 700);
        
        if (style.TryGetValue("font-style", out var fstyle))
            props.Italic = fstyle.Equals("italic", StringComparison.OrdinalIgnoreCase);

        if (style.TryGetValue("text-decoration", out var decor))
            props.Underline = decor.Contains("underline", StringComparison.OrdinalIgnoreCase);

        if (style.TryGetValue("color", out var color))
            props.Color = ParseColor(color);

        if (style.TryGetValue("font-size", out var size))
            props.FontSize = ParseFontSize(size);

        if (style.TryGetValue("font-family", out var family))
            props.FontFamily = family.Split(',')[0].Trim('\'', '\"', ' ');

        // Manual tag check for fallback if StyleResolver is disabled or for specific tags
        var current = node;
        while (current != null)
        {
            if (current.TagName == "b" || current.TagName == "strong") props.Bold = true;
            if (current.TagName == "i" || current.TagName == "em") props.Italic = true;
            if (current.TagName == "u") props.Underline = true;
            current = current.Parent;
        }

        return props;
    }

    private static string? GetTextAlign(HtmlNode node)
    {
        if (node.ComputedStyle.TryGetValue("text-align", out var align))
        {
            return align switch
            {
                "center" => "center",
                "right" => "right",
                "justify" => "both",
                _ => "left"
            };
        }
        return null;
    }

    private static int ParseFontSize(string size, int basePt = 12)
    {
        // Extended parser: pt, px, em, rem, %, and bare numbers.
        // basePt is used for em/rem/% conversions; default is 12pt.
        if (string.IsNullOrWhiteSpace(size)) return 0;
        size = size.Trim().ToLowerInvariant();
        try {
            if (size.EndsWith("pt"))
                return (int)double.Parse(size.Substring(0, size.Length - 2));
            if (size.EndsWith("px"))
                return (int)(double.Parse(size.Substring(0, size.Length - 2)) * 0.75);
            if (size.EndsWith("em") || size.EndsWith("rem"))
            {
                var num = double.Parse(size.Substring(0, size.Length - (size.EndsWith("rem") ? 3 : 2)));
                return (int)(num * basePt);
            }
            if (size.EndsWith("%"))
            {
                var num = double.Parse(size.Substring(0, size.Length - 1));
                return (int)(basePt * num / 100.0);
            }
            if (double.TryParse(size, out var val))
                return (int)val;
        } catch {}
        return 0;
    }

    private static bool ParsePercentage(string s, out double val)
    {
        val = 0;
        s = s.Trim();
        if (s.EndsWith("%"))
        {
            if (double.TryParse(s.Substring(0, s.Length - 1), out var num))
            {
                val = num / 100.0;
                return true;
            }
        }
        return false;
    }

    private static (byte r, byte g, byte b) HslToRgb(double h, double s, double l)
    {
        h = h % 360;
        if (h < 0) h += 360;
        double c = (1 - Math.Abs(2 * l - 1)) * s;
        double x = c * (1 - Math.Abs((h / 60) % 2 - 1));
        double m = l - c/2;
        double r1=0, g1=0, b1=0;
        if (h < 60) { r1=c; g1=x; b1=0; }
        else if (h < 120) { r1=x; g1=c; b1=0; }
        else if (h < 180) { r1=0; g1=c; b1=x; }
        else if (h < 240) { r1=0; g1=x; b1=c; }
        else if (h < 300) { r1=x; g1=0; b1=c; }
        else { r1=c; g1=0; b1=x; }
        byte r = (byte)Math.Round((r1 + m) * 255);
        byte g = (byte)Math.Round((g1 + m) * 255);
        byte b = (byte)Math.Round((b1 + m) * 255);
        return (r,g,b);
    }

    private static string? ParseColor(string color)
    {
        if (string.IsNullOrWhiteSpace(color)) return null;
        color = color.Trim().ToLowerInvariant();
        // hex notation
        if (color.StartsWith("#"))
        {
            // expand shorthand #abc to #aabbcc
            if (color.Length == 4)
            {
                var r = color[1]; var g = color[2]; var b = color[3];
                return $"#{r}{r}{g}{g}{b}{b}";
            }
            // ignore alpha channel if present (#rrggbbaa)
            if (color.Length == 9)
            {
                return color.Substring(0, 7);
            }
            return color;
        }

        if (color.StartsWith("rgb("))
        {
            var inside = color.Substring(4, color.Length - 5);
            var parts = inside.Split(',');
            if (parts.Length == 3 &&
                byte.TryParse(parts[0], out var r) &&
                byte.TryParse(parts[1], out var g) &&
                byte.TryParse(parts[2], out var b))
            {
                return $"#{r:X2}{g:X2}{b:X2}".ToLowerInvariant();
            }
        }
        if (color.StartsWith("rgba("))
        {
            var inside = color.Substring(5, color.Length - 6);
            var parts = inside.Split(',');
            if (parts.Length == 4 &&
                byte.TryParse(parts[0], out var r) &&
                byte.TryParse(parts[1], out var g) &&
                byte.TryParse(parts[2], out var b))
            {
                // ignore alpha
                return $"#{r:X2}{g:X2}{b:X2}".ToLowerInvariant();
            }
        }

        // hsl/hsla support
        if (color.StartsWith("hsl("))
        {
            var inside = color.Substring(4, color.Length - 5);
            var parts = inside.Split(',');
            if (parts.Length == 3 &&
                double.TryParse(parts[0], out var h) &&
                ParsePercentage(parts[1], out var s) &&
                ParsePercentage(parts[2], out var l))
            {
                var (r, g, b) = HslToRgb(h, s, l);
                return $"#{r:X2}{g:X2}{b:X2}".ToLowerInvariant();
            }
        }
        if (color.StartsWith("hsla("))
        {
            var inside = color.Substring(5, color.Length - 6);
            var parts = inside.Split(',');
            if (parts.Length >= 4 &&
                double.TryParse(parts[0], out var h) &&
                ParsePercentage(parts[1], out var s) &&
                ParsePercentage(parts[2], out var l))
            {
                var (r, g, b) = HslToRgb(h, s, l);
                return $"#{r:X2}{g:X2}{b:X2}".ToLowerInvariant();
            }
        }

        if (color == "transparent")
            return null;

        // named colors
        if (_namedColors.TryGetValue(color, out var hex))
            return hex;

        return null;
    }

    private static readonly Dictionary<string, string> _namedColors = new(StringComparer.OrdinalIgnoreCase)
    {
        {"black","#000000"}, {"white","#ffffff"}, {"red","#ff0000"}, {"green","#008000"},
        {"blue","#0000ff"}, {"yellow","#ffff00"}, {"cyan","#00ffff"}, {"magenta","#ff00ff"},
        {"gray","#808080"}, {"silver","#c0c0c0"}, {"maroon","#800000"}, {"olive","#808000"},
        {"purple","#800080"}, {"teal","#008080"}, {"navy","#000080"},
        {"lime","#00ff00"}, {"aqua","#00ffff"}, {"fuchsia","#ff00ff"}, {"orange","#ffa500"},
        {"brown","#a52a2a"}, {"pink","#ffc0cb"}, {"gold","#ffd700"}, {"beige","#f5f5dc"}
    };
    private async Task<ImageData?> ProcessImage(string src)
    {
        try {
            byte[] data;
            string contentType = "image/png";

            if (src.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                var parts = src.Split(',');
                if (parts.Length < 2) return null;
                data = System.Convert.FromBase64String(parts[1]);
                var header = parts[0];
                if (header.Contains("image/jpeg")) contentType = "image/jpeg";
                else if (header.Contains("image/gif")) contentType = "image/gif";
            }
            else if (_options.DownloadImages && (src.StartsWith("http") || src.StartsWith("https")))
            {
                using var cts = new CancellationTokenSource(_options.ImageDownloadTimeout);
                var response = await _httpClient.GetAsync(src, cts.Token);
                if (!response.IsSuccessStatusCode) return null;
                data = await response.Content.ReadAsByteArrayAsync();
                contentType = response.Content.Headers.ContentType?.MediaType ?? "image/png";
            }
            else if (File.Exists(src))
            {
                data = await File.ReadAllBytesAsync(src);
                var ext = Path.GetExtension(src).ToLowerInvariant();
                contentType = ext switch {
                    ".jpg" or ".jpeg" => "image/jpeg",
                    ".gif" => "image/gif",
                    _ => "image/png"
                };
            }
            else return null;

            if (data.Length > _options.MaxImageSize) return null;

            int width = 300, height = 200;
            try
            {
                using var ms2 = new MemoryStream(data);
#pragma warning disable SYSLIB0011 // Image.FromStream is obsolete on some platforms
                using var img = System.Drawing.Image.FromStream(ms2);
                width = img.Width;
                height = img.Height;
#pragma warning restore SYSLIB0011
            }
            catch
            {
                // keep defaults if dimension extraction fails
            }

            return new ImageData { Data = data, ContentType = contentType, Width = width, Height = height };
        } catch { return null; }
    }

    private class ImageData
    {
        public byte[] Data { get; set; } = Array.Empty<byte>();
        public string ContentType { get; set; } = string.Empty;
        public int Width { get; set; }
        public int Height { get; set; }
    }

    private static string GetNumberingXml()
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<w:numbering xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">");

        // abstractNumId=0: bullets
        sb.Append("<w:abstractNum w:abstractNumId=\"0\">");
        sb.Append("<w:multiLevelType w:val=\"hybridMultilevel\"/>");
        for (int lvl = 0; lvl <= MaxNumberingLevel; lvl++)
        {
            int left = GetListLeftIndent(lvl);
            int tabPos = GetListTabPos(lvl);
            sb.Append($"<w:lvl w:ilvl=\"{lvl}\">");
            sb.Append("<w:start w:val=\"1\"/>");
            sb.Append("<w:numFmt w:val=\"bullet\"/>");
            sb.Append($"<w:lvlText w:val=\"{GetBulletChar(lvl)}\"/>");
            sb.Append("<w:suff w:val=\"space\"/>");
            sb.Append("<w:lvlJc w:val=\"left\"/>");
            sb.Append("<w:pPr>");
            sb.Append($"<w:tabs><w:tab w:val=\"num\" w:pos=\"{tabPos}\"/></w:tabs>");
            sb.Append($"<w:ind w:left=\"{left}\" w:hanging=\"{ListHangingIndent}\"/>");
            sb.Append("</w:pPr>");
            var font = GetBulletFont(lvl);
            sb.Append($"<w:rPr><w:rFonts w:ascii=\"{font}\" w:hAnsi=\"{font}\" w:hint=\"default\"/></w:rPr>");
            sb.Append("</w:lvl>");
        }
        sb.Append("</w:abstractNum>");

        // abstractNumId=1: decimal numbering
        sb.Append("<w:abstractNum w:abstractNumId=\"1\">");
        sb.Append("<w:multiLevelType w:val=\"hybridMultilevel\"/>");
        for (int lvl = 0; lvl <= MaxNumberingLevel; lvl++)
        {
            int left = GetListLeftIndent(lvl);
            int tabPos = GetListTabPos(lvl);
            sb.Append($"<w:lvl w:ilvl=\"{lvl}\">");
            sb.Append("<w:start w:val=\"1\"/>");
            sb.Append($"<w:numFmt w:val=\"{GetNumberFmt(lvl)}\"/>");
            sb.Append($"<w:lvlText w:val=\"{BuildLvlText(lvl)}\"/>");
            sb.Append("<w:suff w:val=\"space\"/>");
            sb.Append("<w:lvlJc w:val=\"left\"/>");
            sb.Append("<w:pPr>");
            sb.Append($"<w:tabs><w:tab w:val=\"num\" w:pos=\"{tabPos}\"/></w:tabs>");
            sb.Append($"<w:ind w:left=\"{left}\" w:hanging=\"{ListHangingIndent}\"/>");
            sb.Append("</w:pPr>");
            sb.Append("</w:lvl>");
        }
        sb.Append("</w:abstractNum>");

        // concrete nums
        sb.Append("<w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num>");
        sb.Append("<w:num w:numId=\"2\"><w:abstractNumId w:val=\"1\"/></w:num>");
        sb.Append("</w:numbering>");
        return sb.ToString();
    }

    private static bool IsBlockInsideListItem(HtmlNode node)
    {
        var tag = node.TagName;
        if (tag == "#text") return false;
        if (tag == "br" || tag == "span" || tag == "a" || tag == "b" || tag == "strong" || tag == "i" || tag == "em" || tag == "u" || tag == "font" || tag == "img")
            return false;

        // Lists and tables are block for our purposes; also paragraphs/divs/headings.
        return tag == "ul" || tag == "ol" || tag == "table" || tag == "p" || tag == "div" ||
               tag == "h1" || tag == "h2" || tag == "h3" || tag == "h4" || tag == "h5" || tag == "h6";
    }

    private static bool IsParagraphLikeBlock(HtmlNode node)
    {
        var tag = node.TagName;
        return tag == "p" ||
               tag == "h1" || tag == "h2" || tag == "h3" || tag == "h4" || tag == "h5" || tag == "h6";
    }

    private static bool IsContainerBlock(HtmlNode node)
    {
        // A container that often wraps other blocks inside <li>.
        return node.TagName == "div" || node.TagName == "section" || node.TagName == "article";
    }

    private static bool IsHardBlock(HtmlNode node)
    {
        // Blocks that must be emitted as-is (nested lists, tables).
        return node.TagName == "ul" || node.TagName == "ol" || node.TagName == "table";
    }

    private static void StartParagraphLikeNode(HtmlNode? node, DocumentBuilder builder, int? listNumId = null, int listLevel = 0, int? indentLeft = null, int? indentHanging = null, int? tabPos = null)
    {
        if (node == null)
        {
            builder.StartParagraph(
                style: (listNumId.HasValue || indentHanging.HasValue) ? "ListParagraph" : null,
                listNumId: listNumId,
                listLevel: listLevel,
                indentLeft: indentLeft,
                indentHanging: indentHanging,
                tabPos: tabPos
            );
            return;
        }

        string? style = null;
        if (node.TagName.Length == 2 && (node.TagName[0] == 'h' || node.TagName[0] == 'H') && char.IsDigit(node.TagName[1]))
        {
            int level = int.TryParse(node.TagName.Substring(1), out var l) ? l : 1;
            style = $"Heading{level}";
        }
        else if (listNumId.HasValue || indentHanging.HasValue)
        {
            style = "ListParagraph";
        }

        var (before, after) = GetMarginSpacing(node);
        var (padLeft, padRight) = GetPaddingSpacing(node);
        var bg = node.ComputedStyle != null && node.ComputedStyle.TryGetValue("background-color", out var bgc) ? ParseColor(bgc) : null;
        var textAlign = GetTextAlign(node);

        // When this paragraph is a list marker paragraph, let numbering.xml control indentation;
        // don't mix CSS padding into <w:ind> (it can lead to doubled indentation).
        int? resolvedIndentLeft = indentLeft;
        int? resolvedIndentRight = null;
        if (!listNumId.HasValue)
        {
            resolvedIndentLeft ??= padLeft;
            resolvedIndentRight = padRight;
        }

        builder.StartParagraph(
            style: style,
            textAlign: textAlign,
            listNumId: listNumId,
            listLevel: listLevel,
            spacingBeforeTwips: before,
            spacingAfterTwips: after,
            indentLeft: resolvedIndentLeft,
            indentRight: resolvedIndentRight,
            shdColor: bg,
            indentHanging: indentHanging,
            tabPos: tabPos
        );
    }

    private static string GetNumberFmt(int level)
    {
        // A slightly nicer default than "decimal everywhere", but still stable for Word.
        return level switch
        {
            0 => "decimal",
            1 => "lowerLetter",
            2 => "lowerRoman",
            _ => "decimal"
        };
    }

    private static string BuildLvlText(int level)
    {
        // 0 -> %1. ; 1 -> %1.%2. ; etc.
        var sb = new StringBuilder();
        for (int i = 1; i <= level + 1; i++)
        {
            if (i > 1) sb.Append('.');
            sb.Append('%');
            sb.Append(i);
        }
        sb.Append('.');
        return sb.ToString();
    }

    private static byte[] BuildDocx(string documentXml, string? relationsXml = null, List<(string Name, byte[] Data)>? media = null, string? stylesXml = null, string? fontsXml = null, string? headerXml = null, string? footerXml = null)
    {
        using var zip = new ZipArchiveHelper();
        
        // [Content_Types].xml
        zip.AddEntry("[Content_Types].xml", 
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>" +
            "<Default Extension=\"png\" ContentType=\"image/png\"/>" +
            "<Default Extension=\"gif\" ContentType=\"image/gif\"/>" +
            "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
            "<Override PartName=\"/word/numbering.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>" +
            "<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>" +
            "<Override PartName=\"/word/fontTable.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml\"/>" +
            (headerXml != null ? "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>" : "") +
            (footerXml != null ? "<Override PartName=\"/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\"/>" : "") +
            "</Types>");
        
        // _rels/.rels
        zip.AddEntry("_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
            "</Relationships>");
        
        // word/_rels/document.xml.rels
        var docRels = new StringBuilder();
        docRels.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        docRels.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        docRels.Append("<Relationship Id=\"rIdNumbering\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/>");
        docRels.Append("<Relationship Id=\"rIdStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
        docRels.Append("<Relationship Id=\"rIdFonts\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable\" Target=\"fontTable.xml\"/>");
        
        if (headerXml != null)
            docRels.Append("<Relationship Id=\"rIdHeader\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>");
        if (footerXml != null)
            docRels.Append("<Relationship Id=\"rIdFooter\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer1.xml\"/>");

        if (!string.IsNullOrEmpty(relationsXml))
        {
            // Extract intermediate relationships and append
            var tempRels = relationsXml.Replace("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>", "")
                                     .Replace("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">", "")
                                     .Replace("</Relationships>", "");
            docRels.Append(tempRels);
        }
        docRels.Append("</Relationships>");

        zip.AddEntry("word/_rels/document.xml.rels", docRels.ToString());

        // word/document.xml
        zip.AddEntry("word/document.xml", documentXml);

        // word/numbering.xml
        zip.AddEntry("word/numbering.xml", GetNumberingXml());

        // word/styles.xml
        if (!string.IsNullOrEmpty(stylesXml))
            zip.AddEntry("word/styles.xml", stylesXml);

        // word/fontTable.xml
        if (!string.IsNullOrEmpty(fontsXml))
            zip.AddEntry("word/fontTable.xml", fontsXml);

        // word/header1.xml / footer1.xml
        if (!string.IsNullOrEmpty(headerXml))
            zip.AddEntry("word/header1.xml", headerXml);
        if (!string.IsNullOrEmpty(footerXml))
            zip.AddEntry("word/footer1.xml", footerXml);

        // word/media/
        if (media != null)
        {
            foreach (var m in media)
            {
                zip.AddEntry($"word/media/{m.Name}", m.Data);
            }
        }
        
        return zip.ToArray();
    }

    public void Dispose()
    {
        _httpClient.Dispose();
    }
}
