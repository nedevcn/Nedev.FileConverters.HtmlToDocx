using Nedev.HtmlToDocx.Core.Docx;
using Nedev.HtmlToDocx.Core.Html;

namespace Nedev.HtmlToDocx.Core.Conversion;

public class ConverterOptions
{
    public bool PreserveStyles { get; set; } = true;
    public bool DownloadImages { get; set; } = true;
    public int MaxImageSize { get; set; } = 5 * 1024 * 1024;
    public TimeSpan ImageDownloadTimeout { get; set; } = TimeSpan.FromSeconds(30);
}

public class HtmlToDocxConverter : IDisposable
{
    private readonly ConverterOptions _options;

    public HtmlToDocxConverter(ConverterOptions? options = null)
    {
        _options = options ?? new ConverterOptions();
    }

    public byte[] Convert(string html)
    {
        var document = HtmlDocument.Parse(html);
        var builder = new DocumentBuilder();
        ConvertNode(document.Root, builder);
        return BuildDocx(builder.Build());
    }

    public async Task<byte[]> ConvertAsync(string html, CancellationToken cancellationToken = default)
    {
        return await Task.Run(() => Convert(html), cancellationToken);
    }

    private void ConvertNode(HtmlNode node, DocumentBuilder builder)
    {
        foreach (var child in node.Children)
        {
            switch (child.TagName.ToLowerInvariant())
            {
                case "h1":
                case "h2":
                case "h3":
                case "h4":
                case "h5":
                case "h6":
                    var level = int.Parse(child.TagName.Substring(1));
                    builder.AddHeading(ExtractText(child), level);
                    break;
                case "p":
                    builder.AddParagraph(ExtractText(child));
                    break;
                case "table":
                    ConvertTable(child, builder);
                    break;
                case "div":
                case "section":
                case "article":
                    ConvertNode(child, builder);
                    break;
                case "#text":
                    if (!string.IsNullOrWhiteSpace(child.Text))
                        builder.AddParagraph(child.Text);
                    break;
            }
        }
    }

    private void ConvertTable(HtmlNode node, DocumentBuilder builder)
    {
        var table = new TableData();
        foreach (var rowNode in node.Children.Where(c => c.TagName == "tr"))
        {
            var row = new TableRow();
            foreach (var cellNode in rowNode.Children.Where(c => c.TagName == "td" || c.TagName == "th"))
            {
                row.Cells.Add(new TableCell { Text = ExtractText(cellNode) });
            }
            if (row.Cells.Count > 0)
                table.Rows.Add(row);
        }
        if (table.Rows.Count > 0)
            builder.AddTable(table);
    }

    private static string ExtractText(HtmlNode node)
    {
        if (node.TagName == "#text")
            return node.Text;
        
        var sb = new System.Text.StringBuilder();
        foreach (var child in node.Children)
        {
            if (child.TagName == "#text")
                sb.Append(child.Text);
            else
                sb.Append(ExtractText(child));
        }
        return sb.ToString().Trim();
    }

    private static byte[] BuildDocx(string documentXml)
    {
        using var zip = new ZipArchiveHelper();
        
        // [Content_Types].xml
        zip.AddEntry("[Content_Types].xml", 
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
            "</Types>");
        
        // _rels/.rels
        zip.AddEntry("_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
            "</Relationships>");
        
        // word/_rels/document.xml.rels
        zip.AddEntry("word/_rels/document.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "</Relationships>");
        
        // word/document.xml
        zip.AddEntry("word/document.xml", documentXml);
        
        return zip.ToArray();
    }

    public void Dispose()
    {
    }
}
