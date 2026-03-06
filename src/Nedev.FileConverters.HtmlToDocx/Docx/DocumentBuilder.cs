using System.Text;
using System.Collections.Generic;
using Nedev.FileConverters.HtmlToDocx.Core.Models;

namespace Nedev.FileConverters.HtmlToDocx.Core.Docx;

public sealed class DocumentBuilder
{
    private readonly StringBuilder _bodyXml = new();
    private readonly StringBuilder _headerXml = new();
    private readonly StringBuilder _footerXml = new();
    private StringBuilder _currentXml;
    
    private readonly List<(string Id, string Target, string Type)> _relationships = new();
    private readonly List<(string Name, byte[] Data)> _media = new();
    private bool _inParagraph;
    private int _relCount = 0;
    private int _mediaCount = 0;

    public DocumentBuilder()
    {
        _currentXml = _bodyXml;
    }

    public void SwitchToHeader() => _currentXml = _headerXml;
    public void SwitchToFooter() => _currentXml = _footerXml;
    public void SwitchToBody() => _currentXml = _bodyXml;

    public void StartParagraph(string? style = null, string? textAlign = null, int? listNumId = null, int listLevel = 0)
    {
        if (_inParagraph) EndParagraph();
        
        _currentXml.Append("<w:p>");
        _currentXml.Append("<w:pPr>");
        if (!string.IsNullOrEmpty(style))
            _currentXml.Append($"<w:pStyle w:val=\"{style}\"/>");
        if (!string.IsNullOrEmpty(textAlign))
            _currentXml.Append($"<w:jc w:val=\"{textAlign}\"/>");
        
        if (listNumId.HasValue)
        {
            _currentXml.Append("<w:numPr>");
            _currentXml.Append($"<w:ilvl w:val=\"{listLevel}\"/>");
            _currentXml.Append($"<w:numId w:val=\"{listNumId.Value}\"/>");
            _currentXml.Append("</w:numPr>");
        }

        _currentXml.Append("</w:pPr>");
        _inParagraph = true;
    }

    public void EndParagraph()
    {
        if (_inParagraph)
        {
            _currentXml.Append("</w:p>");
            _inParagraph = false;
        }
    }

    public void AddRun(string text, RunProperties? props = null)
    {
        if (!_inParagraph) StartParagraph();

        _currentXml.Append("<w:r>");
        if (props != null)
        {
            _currentXml.Append("<w:rPr>");
            if (props.Bold) _currentXml.Append("<w:b/>");
            if (props.Italic) _currentXml.Append("<w:i/>");
            if (props.Underline) _currentXml.Append("<w:u w:val=\"single\"/>");
            if (!string.IsNullOrEmpty(props.Color))
                _currentXml.Append($"<w:color w:val=\"{props.Color.TrimStart('#')}\"/>");
            if (props.FontSize > 0)
                _currentXml.Append($"<w:sz w:val=\"{props.FontSize * 2}\"/>"); // Half-points
            if (!string.IsNullOrEmpty(props.FontFamily))
                _currentXml.Append($"<w:rFonts w:ascii=\"{props.FontFamily}\" w:hAnsi=\"{props.FontFamily}\"/>");
            _currentXml.Append("</w:rPr>");
        }
        _currentXml.Append("<w:t xml:space=\"preserve\">");
        _currentXml.Append(EscapeXml(text));
        _currentXml.Append("</w:t></w:r>");
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

        _currentXml.Append("<w:r><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">");
        _currentXml.Append($"<wp:extent cx=\"{widthEmu}\" cy=\"{heightEmu}\"/>");
        _currentXml.Append("<wp:docPr id=\"1\" name=\"Image\"/>");
        _currentXml.Append("<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>");
        _currentXml.Append("<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">");
        _currentXml.Append("<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">");
        _currentXml.Append("<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">");
        _currentXml.Append("<pic:nvPicPr><pic:cNvPr id=\"0\" name=\"Picture\"/><pic:cNvPicPr/></pic:nvPicPr>");
        _currentXml.Append("<pic:blipFill>");
        _currentXml.Append($"<a:blip r:embed=\"{relId}\"/>");
        _currentXml.Append("<a:stretch><a:fillRect/></a:stretch>");
        _currentXml.Append("</pic:blipFill>");
        _currentXml.Append("<pic:spPr>");
        _currentXml.Append("<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" + widthEmu + "\" cy=\"" + heightEmu + "\"/></a:xfrm>");
        _currentXml.Append("<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
        _currentXml.Append("</pic:spPr>");
        _currentXml.Append("</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>");
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

        _currentXml.Append("<w:tbl>");
        _currentXml.Append("<w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/><w:tblBorders>");
        _currentXml.Append("<w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        _currentXml.Append("<w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        _currentXml.Append("<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        _currentXml.Append("<w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        _currentXml.Append("<w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        _currentXml.Append("<w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"auto\"/>");
        _currentXml.Append("</w:tblBorders></w:tblPr>");

        foreach (var row in table.Rows)
        {
            _currentXml.Append("<w:tr>");
            foreach (var cell in row.Cells)
            {
                _currentXml.Append("<w:tc>");
                _currentXml.Append("<w:tcPr>");
                if (cell.ColSpan > 1)
                    _bodyXml.Append($"<w:gridSpan w:val=\"{cell.ColSpan}\"/>");
                if (cell.RowMerge == RowMergeType.Restart)
                    _bodyXml.Append("<w:vMerge w:val=\"restart\"/>");
                else if (cell.RowMerge == RowMergeType.Continue)
                    _bodyXml.Append("<w:vMerge/>");
                _currentXml.Append("</w:tcPr>");

                // For now, tables contain simple text paragraphs
                _currentXml.Append("<w:p><w:r><w:t xml:space=\"preserve\">");
                _currentXml.Append(EscapeXml(cell.Text));
                _currentXml.Append("</w:t></w:r></w:p>");
                _currentXml.Append("</w:tc>");
            }
            _currentXml.Append("</w:tr>");
        }
        _currentXml.Append("</w:tbl>");
    }

    public void AddPageNumberField()
    {
        if (!_inParagraph) StartParagraph();
        _currentXml.Append("<w:fldSimple w:instr=\" PAGE \"/>");
    }

    public void AddTotalPagesField()
    {
        if (!_inParagraph) StartParagraph();
        _currentXml.Append("<w:fldSimple w:instr=\" NUMPAGES \"/>");
    }

    public void AddTableOfContents(int levels = 3)
    {
        if (_inParagraph) EndParagraph();
        _currentXml.Append("<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val=\"Table of Contents\"/><w:docPartUnique/></w:docPartObj></w:sdtPr>");
        _currentXml.Append("<w:sdtContent>");
        _currentXml.Append($"<w:p><w:pPr><w:pStyle w:val=\"TOCHeading\"/></w:pPr><w:r><w:t>Table of Contents</w:t></w:r></w:p>");
        _currentXml.Append($"<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/><w:instrText xml:space=\"preserve\"> TOC \\o \"1-{levels}\" \\h \\z \\u </w:instrText><w:fldChar w:fldCharType=\"separate\"/></w:r>");
        _currentXml.Append("<w:r><w:t>Updating...</w:t></w:r>");
        _currentXml.Append("<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>");
        _currentXml.Append("</w:sdtContent></w:sdt>");
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

        sb.Append("<w:sectPr>");
        if (_headerXml.Length > 0)
            sb.Append("<w:headerReference w:type=\"default\" r:id=\"rIdHeader\"/>");
        if (_footerXml.Length > 0)
            sb.Append("<w:footerReference w:type=\"default\" r:id=\"rIdFooter\"/>");
            
        sb.Append($"<w:pgSz w:w=\"{w}\" w:h=\"{h}\"/><w:pgMar w:top=\"{m.Top}\" w:right=\"{m.Right}\" w:bottom=\"{m.Bottom}\" w:left=\"{m.Left}\"/></w:sectPr>");
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

    public string BuildHeader() => BuildPart(_headerXml, "hdr");
    public string BuildFooter() => BuildPart(_footerXml, "ftr");

    private string BuildPart(StringBuilder content, string rootTag)
    {
        return $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
               $"<w:{rootTag} xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
               $"xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
               $"xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">" +
               $"{content}" +
               $"</w:{rootTag}>";
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
