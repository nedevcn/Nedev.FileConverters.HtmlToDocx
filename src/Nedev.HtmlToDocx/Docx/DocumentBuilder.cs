using System.Text;

namespace Nedev.HtmlToDocx.Core.Docx;

public class DocumentBuilder
{
    private readonly StringBuilder _documentXml = new();
    private int _paragraphCount = 0;

    public void AddParagraph(string text, bool isBold = false, bool isItalic = false)
    {
        _documentXml.Append("<w:p>");
        _documentXml.Append("<w:pPr><w:pStyle w:val=\"Normal\"/></w:pPr>");
        _documentXml.Append("<w:r>");
        
        if (isBold || isItalic)
        {
            _documentXml.Append("<w:rPr>");
            if (isBold) _documentXml.Append("<w:b/>");
            if (isItalic) _documentXml.Append("<w:i/>");
            _documentXml.Append("</w:rPr>");
        }
        
        _documentXml.Append("<w:t>");
        _documentXml.Append(EscapeXml(text));
        _documentXml.Append("</w:t></w:r></w:p>");
        _paragraphCount++;
    }

    public void AddHeading(string text, int level)
    {
        _documentXml.Append($"<w:p><w:pPr><w:pStyle w:val=\"Heading{level}\"/></w:pPr>");
        _documentXml.Append("<w:r><w:t>");
        _documentXml.Append(EscapeXml(text));
        _documentXml.Append("</w:t></w:r></w:p>");
        _paragraphCount++;
    }

    public void AddTable(TableData table)
    {
        _documentXml.Append("<w:tbl>");
        foreach (var row in table.Rows)
        {
            _documentXml.Append("<w:tr>");
            foreach (var cell in row.Cells)
            {
                _documentXml.Append("<w:tc>");
                _documentXml.Append("<w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>");
                _documentXml.Append("<w:p><w:r><w:t>");
                _documentXml.Append(EscapeXml(cell.Text));
                _documentXml.Append("</w:t></w:r></w:p>");
                _documentXml.Append("</w:tc>");
            }
            _documentXml.Append("</w:tr>");
        }
        _documentXml.Append("</w:tbl>");
    }

    public string Build()
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">");
        sb.Append("<w:body>");
        sb.Append(_documentXml);
        sb.Append("<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\"/></w:sectPr>");
        sb.Append("</w:body></w:document>");
        return sb.ToString();
    }

    private static string EscapeXml(string text)
    {
        return text.Replace("&", "&amp;")
                   .Replace("<", "&lt;")
                   .Replace(">", "&gt;")
                   .Replace("\"", "&quot;")
                   .Replace("'", "&apos;");
    }
}

public class TableData
{
    public List<TableRow> Rows { get; } = new();
}

public class TableRow
{
    public List<TableCell> Cells { get; } = new();
}

public class TableCell
{
    public string Text { get; set; } = string.Empty;
}
