using Xunit;
using Nedev.FileConverters.HtmlToDocx.Core.Conversion;
using Nedev.FileConverters.HtmlToDocx.Core.Models;
using System.IO.Compression;

namespace Nedev.FileConverters.HtmlToDocx.Tests
{
    public class TableCellStyleTests
    {
        [Fact]
        public void CellBackgroundAndPaddingAreApplied()
        {
            string html = "<table><tr><td style=\"background-color:rgb(255,200,0);padding:2px 4px\">X</td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string xml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                var entry = zip.GetEntry("word/document.xml");
                using var r=new System.IO.StreamReader(entry.Open());
                xml=r.ReadToEnd();
            }
            // padding left 4px -> 3pt -> 60 twips; right 4px same
            Assert.Contains("w:shd", xml);
            Assert.Contains("fill=\"ffc800\"", xml); // rgb(255,200,0)
            Assert.Contains("w:tcMar", xml);
            Assert.Contains("w:left=\"60\"", xml);
            Assert.Contains("w:right=\"60\"", xml);
        }

        [Fact]
        public void CellInlineFormattingIsPreserved()
        {
            string html = "<table><tr><td><strong>A</strong><span style=\"color:#ff0000\">B</span></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string xml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                var entry = zip.GetEntry("word/document.xml");
                using var r=new System.IO.StreamReader(entry.Open());
                xml=r.ReadToEnd();
            }

            Assert.Contains("<w:b/>", xml);
            Assert.Contains("<w:color w:val=\"ff0000\"/>", xml);
            Assert.Contains(">A</w:t>", xml);
            Assert.Contains(">B</w:t>", xml);
        }

        [Fact]
        public void CellParagraphBlocksArePreserved()
        {
            string html = "<table><tr><td><p>One</p><p>Two</p></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string xml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                var entry = zip.GetEntry("word/document.xml");
                using var r=new System.IO.StreamReader(entry.Open());
                xml=r.ReadToEnd();
            }

            Assert.Contains("One", xml);
            Assert.Contains("Two", xml);
            Assert.True(System.Text.RegularExpressions.Regex.Matches(xml, "<w:tc>.*?<w:p>.*?One.*?</w:p><w:p>.*?Two.*?</w:p>.*?</w:tc>", System.Text.RegularExpressions.RegexOptions.Singleline).Count > 0);
        }

        [Fact]
        public void CellHyperlinkIsRenderedAsRelationshipHyperlink()
        {
            string html = "<table><tr><td><a href=\"https://example.com\">Go</a></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string xml;
            string rels;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                var entry = zip.GetEntry("word/document.xml");
                using var r=new System.IO.StreamReader(entry.Open());
                xml=r.ReadToEnd();
                using var rr = new System.IO.StreamReader(zip.GetEntry("word/_rels/document.xml.rels")!.Open());
                rels = rr.ReadToEnd();
            }

            Assert.Contains("<w:hyperlink r:id=\"rId", xml);
            Assert.Contains(">Go</w:t>", xml);
            Assert.Contains("relationships/hyperlink", rels);
            Assert.Contains("TargetMode=\"External\"", rels);
            Assert.Contains("https://example.com", rels);
        }

        [Fact]
        public void CellImageWithoutEmbed_PreservesAltTextFallback()
        {
            string html = "<table><tr><td><img src=\"https://img.test/a.png\" alt=\"Logo\"/></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string xml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                var entry = zip.GetEntry("word/document.xml");
                using var r=new System.IO.StreamReader(entry.Open());
                xml=r.ReadToEnd();
            }

            Assert.Contains(">Logo</w:t>", xml);
        }

        [Fact]
        public void CellDataUriImage_IsEmbeddedIntoDocxMedia()
        {
            const string base64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2n0AAAAASUVORK5CYII=";
            string html = "<table><tr><td><img src=\"data:image/png;base64," + base64 + "\"/></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);

            string docXml;
            string relsXml;
            bool hasMediaFile;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using (var r = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open()))
                    docXml = r.ReadToEnd();
                using (var r = new System.IO.StreamReader(zip.GetEntry("word/_rels/document.xml.rels")!.Open()))
                    relsXml = r.ReadToEnd();
                hasMediaFile = zip.GetEntry("word/media/image1.png") != null;
            }

            Assert.Contains("<w:drawing>", docXml);
            Assert.Contains("relationships/image", relsXml);
            Assert.True(hasMediaFile);
        }

        [Fact]
        public void CellList_PreservesNumberingStructure()
        {
            string html = "<table><tr><td><ul><li>One</li><li>Two<ol><li>Sub</li></ol></li></ul></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string xml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                var entry = zip.GetEntry("word/document.xml");
                using var r=new System.IO.StreamReader(entry.Open());
                xml=r.ReadToEnd();
            }

            Assert.Contains("One", xml);
            Assert.Contains("Two", xml);
            Assert.Contains("Sub", xml);
            Assert.Contains("<w:numId w:val=\"1\"/>", xml);
            Assert.Contains("<w:numId w:val=\"2\"/>", xml);
            Assert.Contains("<w:ilvl w:val=\"1\"/>", xml);
        }

        [Fact]
        public void CellNestedTable_IsRenderedAsNestedTableXml()
        {
            string html = "<table><tr><td>Outer<table><tr><td>Inner</td></tr></table></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string xml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                var entry = zip.GetEntry("word/document.xml");
                using var r=new System.IO.StreamReader(entry.Open());
                xml=r.ReadToEnd();
            }

            Assert.Contains("Outer", xml);
            Assert.Contains("Inner", xml);
            Assert.True(System.Text.RegularExpressions.Regex.Matches(xml, "<w:tc>[\\s\\S]*<w:tbl>[\\s\\S]*Inner[\\s\\S]*</w:tbl>[\\s\\S]*</w:tc>").Count > 0);
        }

        [Fact]
        public void CellHyperlinkWrappingImage_PreservesHyperlinkAndImageRelationship()
        {
            const string base64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2n0AAAAASUVORK5CYII=";
            string html = "<table><tr><td><a href=\"https://example.com\"><img src=\"data:image/png;base64," + base64 + "\" alt=\"Logo\"/></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string docXml;
            string relsXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
                using var r = new System.IO.StreamReader(zip.GetEntry("word/_rels/document.xml.rels")!.Open());
                relsXml = r.ReadToEnd();
            }

            Assert.Contains("<w:hyperlink r:id=\"rId", docXml);
            Assert.Contains("<w:drawing>", docXml);
            Assert.Contains("relationships/hyperlink", relsXml);
            Assert.Contains("https://example.com", relsXml);
            Assert.Contains("relationships/image", relsXml);
        }

        [Fact]
        public void CellHyperlinkWithTextAndImage_UsesSingleHyperlinkContainer()
        {
            const string base64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2n0AAAAASUVORK5CYII=";
            string html = "<table><tr><td><a href=\"https://example.com\">Go<img src=\"data:image/png;base64," + base64 + "\" alt=\"Logo\"/></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string docXml;
            string relsXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
                using var r = new System.IO.StreamReader(zip.GetEntry("word/_rels/document.xml.rels")!.Open());
                relsXml = r.ReadToEnd();
            }

            Assert.Contains(">Go</w:t>", docXml);
            Assert.True(System.Text.RegularExpressions.Regex.Matches(docXml, "<w:hyperlink[^>]*>[\\s\\S]*>Go</w:t>[\\s\\S]*<w:drawing>[\\s\\S]*</w:hyperlink>").Count > 0);
            Assert.Contains("relationships/hyperlink", relsXml);
            Assert.Contains("relationships/image", relsXml);
        }

        [Fact]
        public void CellHyperlinkInlineSegments_PreservePerRunStyles()
        {
            string html = "<table><tr><td><a href=\"https://example.com\"><span style=\"color:#ff0000\">R</span><strong>B</strong></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string docXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
            }

            Assert.True(System.Text.RegularExpressions.Regex.Matches(docXml, "<w:hyperlink[^>]*>[\\s\\S]*>R</w:t>[\\s\\S]*>B</w:t>[\\s\\S]*</w:hyperlink>").Count > 0);
            Assert.Contains("<w:color w:val=\"ff0000\"/>", docXml);
            Assert.Contains("<w:b/>", docXml);
        }

        [Fact]
        public void CellHyperlink_WithBr_PreservesLineBreakInHyperlink()
        {
            string html = "<table><tr><td><a href=\"https://example.com\">A<br/>B</a></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string docXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
            }

            Assert.True(System.Text.RegularExpressions.Regex.Matches(docXml, "<w:hyperlink[^>]*>[\\s\\S]*>A</w:t>[\\s\\S]*<w:br/>[\\s\\S]*>B</w:t>[\\s\\S]*</w:hyperlink>").Count > 0);
        }

        [Fact]
        public void CellHyperlink_BlockContainers_InsertBreakBetweenBlocks()
        {
            string html = "<table><tr><td><a href=\"https://example.com\"><div>Top</div><div>Bottom</div></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string docXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
            }

            Assert.True(System.Text.RegularExpressions.Regex.Matches(docXml, "<w:hyperlink[^>]*>[\\s\\S]*>Top</w:t>[\\s\\S]*<w:br/>[\\s\\S]*>Bottom</w:t>[\\s\\S]*</w:hyperlink>").Count > 0);
        }

        [Fact]
        public void CellHyperlink_ListItems_AreRenderedAsItemRuns()
        {
            string html = "<table><tr><td><a href=\"https://example.com\"><ul><li>One</li><li>Two<ol><li>Sub</li></ol></li></ul></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string docXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
            }

            Assert.True(System.Text.RegularExpressions.Regex.Matches(docXml, "<w:hyperlink[^>]*>[\\s\\S]*-\\s*</w:t>[\\s\\S]*>One</w:t>[\\s\\S]*-\\s*</w:t>[\\s\\S]*>Two</w:t>[\\s\\S]*(1|a)\\.\\s*</w:t>[\\s\\S]*>Sub</w:t>[\\s\\S]*</w:hyperlink>").Count > 0);
            Assert.Contains("<w:br/>", docXml);
        }

        [Fact]
        public void CellHyperlink_OrderedList_NestedUsesAlphaMarker()
        {
            string html = "<table><tr><td><a href=\"https://example.com\"><ol><li>Top<ol><li>Sub</li></ol></li></ol></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string docXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
            }

            Assert.True(System.Text.RegularExpressions.Regex.Matches(docXml, "<w:hyperlink[^>]*>[\\s\\S]*1\\.\\s*</w:t>[\\s\\S]*>Top</w:t>[\\s\\S]*a\\.\\s*</w:t>[\\s\\S]*>Sub</w:t>[\\s\\S]*</w:hyperlink>").Count > 0);
        }

        [Fact]
        public void CellHyperlink_OrderedMarkers_ArePaddedForAlignment()
        {
            string html = "<table><tr><td><a href=\"https://example.com\"><ol><li>One</li><li>Two</li></ol></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string docXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
            }

            Assert.True(System.Text.RegularExpressions.Regex.Matches(docXml, "1\\.\\s*</w:t>[\\s\\S]*>One</w:t>").Count > 0);
            Assert.True(System.Text.RegularExpressions.Regex.Matches(docXml, "2\\.\\s*</w:t>[\\s\\S]*>Two</w:t>").Count > 0);
        }

        [Fact]
        public void CellHyperlink_ListItemComplexInline_PreservesStylesAndImage()
        {
            const string base64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2n0AAAAASUVORK5CYII=";
            string html = "<table><tr><td><a href=\"https://example.com\"><ul><li><span style=\"color:#ff0000\">R</span><strong>B</strong><img src=\"data:image/png;base64," + base64 + "\" alt=\"I\"/></li></ul></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string docXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
            }

            Assert.Contains("<w:color w:val=\"ff0000\"/>", docXml);
            Assert.Contains("<w:b/>", docXml);
            Assert.Contains("<w:drawing>", docXml);
            Assert.True(System.Text.RegularExpressions.Regex.Matches(docXml, "<w:hyperlink[^>]*>[\\s\\S]*-\\s*</w:t>[\\s\\S]*>R</w:t>[\\s\\S]*>B</w:t>[\\s\\S]*<w:drawing>[\\s\\S]*</w:hyperlink>").Count > 0);
        }

        [Fact]
        public void CellHyperlink_ListItemColorPriority_PreservesInlineColorOverLinkDefault()
        {
            string html = "<table><tr><td><a href=\"https://example.com\"><ul><li><span style=\"color:#00aa00\">Green</span> Default</li></ul></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string docXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
            }

            Assert.Contains("<w:color w:val=\"00aa00\"/>", docXml);
            Assert.Contains("<w:color w:val=\"0000ff\"/>", docXml);
            Assert.True(System.Text.RegularExpressions.Regex.Matches(docXml, "<w:hyperlink[^>]*>[\\s\\S]*>Green</w:t>[\\s\\S]*> Default</w:t>[\\s\\S]*</w:hyperlink>").Count > 0);
        }

        [Fact]
        public void CellHyperlink_ListMarker_InheritsLiFontSizeAndFamily()
        {
            string html = "<table><tr><td><a href=\"https://example.com\"><ol><li style=\"font-size:16pt;font-family:Calibri\">One</li></ol></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string docXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
            }

            var markerRun = System.Text.RegularExpressions.Regex.Match(docXml, "<w:r>[\\s\\S]*?<w:t xml:space=\"preserve\">1\\.\\s*</w:t>[\\s\\S]*?</w:r>");
            Assert.True(markerRun.Success);
            Assert.Contains("w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\"", markerRun.Value);
            Assert.Contains("w:sz w:val=\"32\"", markerRun.Value);
        }

        [Fact]
        public void CellHyperlink_ListMarker_DefaultUsesLinkBlueColor()
        {
            string html = "<table><tr><td><a href=\"https://example.com\"><ol><li style=\"color:#00aa00\">One</li></ol></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string docXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
            }

            var markerRun = System.Text.RegularExpressions.Regex.Match(docXml, "<w:r>[\\s\\S]*?<w:t xml:space=\"preserve\">1\\.\\s*</w:t>[\\s\\S]*?</w:r>");
            Assert.True(markerRun.Success);
            Assert.Contains("w:color w:val=\"0000ff\"", markerRun.Value);
        }

        [Fact]
        public void CellHyperlink_ListMarker_CanUseLiColorWhenConfigured()
        {
            string html = "<table><tr><td><a href=\"https://example.com\"><ol><li style=\"color:#00aa00\">One</li></ol></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter(new ConverterOptions { UseLinkColorForListMarkersInHyperlinks = false });
            var bytes = converter.Convert(html);
            string docXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
            }

            var markerRun = System.Text.RegularExpressions.Regex.Match(docXml, "<w:r>[\\s\\S]*?<w:t xml:space=\"preserve\">1\\.\\s*</w:t>[\\s\\S]*?</w:r>");
            Assert.True(markerRun.Success);
            Assert.Contains("w:color w:val=\"00aa00\"", markerRun.Value);
        }

        [Fact]
        public void CellHyperlink_ListMarker_DefaultIsUnderlined()
        {
            string html = "<table><tr><td><a href=\"https://example.com\"><ol><li>One</li></ol></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);
            string docXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
            }

            var markerRun = System.Text.RegularExpressions.Regex.Match(docXml, "<w:r>[\\s\\S]*?<w:t xml:space=\"preserve\">1\\.\\s*</w:t>[\\s\\S]*?</w:r>");
            Assert.True(markerRun.Success);
            Assert.Contains("<w:u w:val=\"single\"/>", markerRun.Value);
        }

        [Fact]
        public void CellHyperlink_ListMarker_CanDisableUnderlineWhenConfigured()
        {
            string html = "<table><tr><td><a href=\"https://example.com\"><ol><li>One</li></ol></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter(new ConverterOptions { UnderlineListMarkersInHyperlinks = false });
            var bytes = converter.Convert(html);
            string docXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
            }

            var markerRun = System.Text.RegularExpressions.Regex.Match(docXml, "<w:r>[\\s\\S]*?<w:t xml:space=\"preserve\">1\\.\\s*</w:t>[\\s\\S]*?</w:r>");
            Assert.True(markerRun.Success);
            Assert.DoesNotContain("<w:u w:val=\"single\"/>", markerRun.Value);
        }

        [Fact]
        public void CellHyperlink_ListMarker_StrategyCanOverrideLegacyBooleans()
        {
            string html = "<table><tr><td><a href=\"https://example.com\"><ol><li style=\"color:#00aa00\">One</li></ol></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter(new ConverterOptions
            {
                UseLinkColorForListMarkersInHyperlinks = false,
                UnderlineListMarkersInHyperlinks = false,
                HyperlinkListMarkerStyleStrategy = HyperlinkListMarkerStyleStrategy.LinkColorAndUnderline
            });
            var bytes = converter.Convert(html);
            string docXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
            }

            var markerRun = System.Text.RegularExpressions.Regex.Match(docXml, "<w:r>[\\s\\S]*?<w:t xml:space=\"preserve\">1\\.\\s*</w:t>[\\s\\S]*?</w:r>");
            Assert.True(markerRun.Success);
            Assert.Contains("w:color w:val=\"0000ff\"", markerRun.Value);
            Assert.Contains("<w:u w:val=\"single\"/>", markerRun.Value);
        }

        [Fact]
        public void CellHyperlink_ListMarker_StrategyInheritColorNoUnderline_Works()
        {
            string html = "<table><tr><td><a href=\"https://example.com\"><ol><li style=\"color:#00aa00\">One</li></ol></a></td></tr></table>";
            using var converter = new HtmlToDocxConverter(new ConverterOptions
            {
                HyperlinkListMarkerStyleStrategy = HyperlinkListMarkerStyleStrategy.InheritColorNoUnderline
            });
            var bytes = converter.Convert(html);
            string docXml;
            using(var ms=new System.IO.MemoryStream(bytes))
            using(var zip=new ZipArchive(ms, ZipArchiveMode.Read, true))
            {
                using var d = new System.IO.StreamReader(zip.GetEntry("word/document.xml")!.Open());
                docXml = d.ReadToEnd();
            }

            var markerRun = System.Text.RegularExpressions.Regex.Match(docXml, "<w:r>[\\s\\S]*?<w:t xml:space=\"preserve\">1\\.\\s*</w:t>[\\s\\S]*?</w:r>");
            Assert.True(markerRun.Success);
            Assert.Contains("w:color w:val=\"00aa00\"", markerRun.Value);
            Assert.DoesNotContain("<w:u w:val=\"single\"/>", markerRun.Value);
        }
    }
}
