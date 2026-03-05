using System.Text;
using System.Collections.Generic;
using Nedev.HtmlToDocx.Core.Models;

namespace Nedev.HtmlToDocx.Core.Docx;

public sealed class DocumentBuilder
{
    private readonly StringBuilder _bodyXml = new();
    private readonly List<(string Id, string Target, string Type)> _relationships = new();
    private readonly List<(string Name, byte[] Data)> _media = new();
    private bool _inParagraph;
    private int _relCount = 0;
    private int _mediaCount = 0;

    public void StartParagraph(string? style = null, string? textAlign = null, int? listNumId = null, int listLevel = 0)
    {
        if (_inParagraph) EndParagraph();
        
        _bodyXml.Append("<w:p>");
        _bodyXml.Append("<w:pPr>");
        if (!string.IsNullOrEmpty(style))
            _bodyXml.Append($"<w:pStyle w:val=\"{style}\"/>");
        if (!string.IsNullOrEmpty(textAlign))
            _bodyXml.Append($"<w:jc w:val=\"{textAlign}\"/>");
        
        if (listNumId.HasValue)
        {
            _bodyXml.Append("<w:numPr>");
            _bodyXml.Append($"<w:ilvl w:val=\"{listLevel}\"/>");
            _bodyXml.Append($"<w:numId w:val=\"{listNumId.Value}\"/>");
            _bodyXml.Append("</w:numPr>");
        }

        _bodyXml.Append("</w:pPr>");
        _inParagraph = true;
    }

    public void EndParagraph()
    {
        if (_inParagraph)
        {
            _bodyXml.Append("</w:p>");
            _inParagraph = false;
        }
    }

    public void AddRun(string text, RunProperties? props = null)
    {
        if (!_inParagraph) StartParagraph();

        _bodyXml.Append("<w:r>");
        if (props != null)
        {
            _bodyXml.Append("<w:rPr>");
            if (props.Bold) _bodyXml.Append("<w:b/>");
            if (props.Italic) _bodyXml.Append("<w:i/>");
            if (props.Underline) _bodyXml.Append("<w:u w:val=\"single\"/>");
            if (!string.IsNullOrEmpty(props.Color))
                _bodyXml.Append($"<w:color w:val=\"{props.Color.TrimStart('#')}\"/>");
            if (props.FontSize > 0)
                _bodyXml.Append($"<w:sz w:val=\"{props.FontSize * 2}\"/>"); // Half-points
            if (!string.IsNullOrEmpty(props.FontFamily))
                _bodyXml.Append($"<w:rFonts w:ascii=\"{props.FontFamily}\" w:hAnsi=\"{props.FontFamily}\"/>");
            _bodyXml.Append("</w:rPr>");
        }
        _bodyXml.Append("<w:t xml:space=\"preserve\">");
        _bodyXml.Append(EscapeXml(text));
        _bodyXml.Append("</w:t></w:r>");
    }

    public void AddHyperlink(string text, string url, RunProperties? props = null)
    {
        if (!_inParagraph) StartParagraph();

        var relId = $"rId{++_relCount}";
        _relationships.Add((relId, url, "hyperlink"));

        _bodyXml.Append($"<w:hyperlink r:id=\"{relId}\">");
        AddRun(text, props ?? new RunProperties { Color = "0000FF", Underline = true });
        _bodyXml.Append("</w:hyperlink>");
    }

    public void AddImage(byte[] data, string contentType, int widthPx = 300, int heightPx = 200)
    {
        if (!_inParagraph) StartParagraph();

        var ext = contentType.Split('/').Last();
        if (ext == "jpeg") ext = "jpg";
        var fileName = $"image{++_mediaCount}.{ext}";
        var relId = $"rId{++_relCount}";
        
        _media.Add((fileName, data));
        _relationships.Add((relId, $"media/{fileName}", "image"));

        // EMUs (English Metric Units): 1 px ~= 9525 EMUs (at 96 DPI)
        long widthEmu = (long)widthPx * 9525;
        long heightEmu = (long)heightPx * 9525;

        _bodyXml.Append("<w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">");
        _bodyXml.Append($"<wp:extent cx=\"{widthEmu}\" cy=\"{heightEmu}\"/>");
        _bodyXml.Append("<wp:docPr id=\"1\" name=\"Image\"/>");
        _bodyXml.Append("<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>");
        _bodyXml.Append("<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">");
        _bodyXml.Append("<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">");
        _bodyXml.Append("<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">");
        _bodyXml.Append("<pic:nvPicPr><pic:cNvPr id=\"0\" name=\"Picture\"/><pic:cNvPicPr/></pic:nvPicPr>");
        _bodyXml.Append("<pic:blipFill>");
        _bodyXml.Append($"<a:blip r:embed=\"{relId}\"/>");
        _bodyXml.Append("<a:stretch><a:fillRect/></a:stretch>");
        _bodyXml.Append("</pic:blipFill>");
        _bodyXml.Append("<pic:spPr>");
        _bodyXml.Append("<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" + widthEmu + "\" cy=\"" + heightEmu + "\"/></a:xfrm>");
        _bodyXml.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
        _bodyXml.Append("</pic:spPr>");
        _bodyXml.Append("</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>");
    }

    public void AddHeading(string text, int level)
    {
        StartParagraph($"Heading{level}");
        AddRun(text);
        EndParagraph();
    }

    public void AddTable(TableData table)
    {
        if (_inParagraph) EndParagraph();

        _bodyXml.Append("<w:tbl>");
        _bodyXml.Append("<w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/><w:tblBorders>");
        _bodyXml.Append("<w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        _bodyXml.Append("<w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        _bodyXml.Append("<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        _bodyXml.Append("<w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        _bodyXml.Append("<w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        _bodyXml.Append("<w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        _bodyXml.Append("</w:tblBorders></w:tblPr>");

        foreach (var row in table.Rows)
        {
            _bodyXml.Append("<w:tr>");
            foreach (var cell in row.Cells)
            {
                _bodyXml.Append("<w:tc>");
                _bodyXml.Append("<w:tcPr>");
                if (cell.ColSpan > 1)
                    _bodyXml.Append($"<w:gridSpan w:val=\"{cell.ColSpan}\"/>");
                if (cell.RowMerge == RowMergeType.Restart)
                    _bodyXml.Append("<w:vMerge w:val=\"restart\"/>");
                else if (cell.RowMerge == RowMergeType.Continue)
                    _bodyXml.Append("<w:vMerge/>");
                _bodyXml.Append("</w:tcPr>");

                // For now, tables contain simple text paragraphs
                _bodyXml.Append("<w:p><w:r><w:t xml:space=\"preserve\">");
                _bodyXml.Append(EscapeXml(cell.Text));
                _bodyXml.Append("</w:t></w:r></w:p>");
                _bodyXml.Append("</w:tc>");
            }
            _bodyXml.Append("</w:tr>");
        }
        _bodyXml.Append("</w:tbl>");
    }

    public string Build(ConverterOptions? options = null)
    {
        if (_inParagraph) EndParagraph();

        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" ");
        sb.Append("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ");
        sb.Append("xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">");
        sb.Append("<w:body>");
        sb.Append(_bodyXml);
        
        var w = options?.PageWidth ?? 11906;
        var h = options?.PageHeight ?? 16838;
        var m = options?.Margins ?? new PageMargins();

        sb.Append($"<w:sectPr><w:pgSz w:w=\"{w}\" w:h=\"{h}\"/><w:pgMar w:top=\"{m.Top}\" w:right=\"{m.Right}\" w:bottom=\"{m.Bottom}\" w:left=\"{m.Left}\"/></w:sectPr>");
        sb.Append("</w:body></w:document>");
        return sb.ToString();
    }

    public string BuildRelationships()
    {
        var sb = new StringBuilder();
        sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
        foreach (var rel in _relationships)
        {
            var typeUrl = rel.Type == "hyperlink" 
                ? "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                : "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
            
            var targetMode = rel.Type == "hyperlink" ? " TargetMode=\"External\"" : "";
            sb.Append($"<Relationship Id=\"{rel.Id}\" Type=\"{typeUrl}\" Target=\"{EscapeXml(rel.Target)}\"{targetMode}/>");
        }
        sb.Append("</Relationships>");
        return sb.ToString();
    }

    public List<(string Name, byte[] Data)> BuildMedia() => _media;

    public string BuildStyles()
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
               "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
               "<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"/><w:sz w:val=\"22\"/></w:rPr></w:rPrDefault></w:docDefaults>" +
               "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:qFormat/></w:style>" +
               "<w:style w:type=\"paragraph\" w:styleId=\"Heading1\"><w:name w:val=\"heading 1\"/><w:next w:val=\"Normal\"/><w:qFormat/><w:pPr><w:outlineLvl w:val=\"0\"/></w:pPr><w:rPr><w:b/><w:sz w:val=\"32\"/></w:rPr></w:style>" +
               "<w:style w:type=\"paragraph\" w:styleId=\"Heading2\"><w:name w:val=\"heading 2\"/><w:next w:val=\"Normal\"/><w:qFormat/><w:pPr><w:outlineLvl w:val=\"1\"/></w:pPr><w:rPr><w:b/><w:sz w:val=\"28\"/></w:rPr></w:style>" +
               "<w:style w:type=\"paragraph\" w:styleId=\"Heading3\"><w:name w:val=\"heading 3\"/><w:next w:val=\"Normal\"/><w:qFormat/><w:pPr><w:outlineLvl w:val=\"2\"/></w:pPr><w:rPr><w:b/><w:sz w:val=\"24\"/></w:rPr></w:style>" +
               "</w:styles>";
    }

    public string BuildFonts()
    {
        return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
               "<w:fontTable xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
               "<w:font w:name=\"Calibri\"><w:panose1 w:val=\"020F0502020204030204\"/><w:charset w:val=\"00\"/><w:family w:val=\"swiss\"/><w:pitch w:val=\"variable\"/></w:font>" +
               "</w:fontTable>";
    }

    private static string EscapeXml(string text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        return text.Replace("&", "&amp;")
                   .Replace("<", "&lt;")
                   .Replace(">", "&gt;")
                   .Replace("\"", "&quot;")
                   .Replace("'", "&apos;");
    }
}
