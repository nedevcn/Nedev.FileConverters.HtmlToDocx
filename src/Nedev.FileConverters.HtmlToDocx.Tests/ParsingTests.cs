using Xunit;
using System;
using System.Reflection;
using Nedev.FileConverters.HtmlToDocx.Core.Conversion;

namespace Nedev.FileConverters.HtmlToDocx.Tests
{
    public class ParsingTests
    {
        private static MethodInfo GetStatic(string name)
            => typeof(HtmlToDocxConverter).GetMethod(name, BindingFlags.NonPublic | BindingFlags.Static)!;

        [Theory]
        [InlineData("12pt", 12)]
        [InlineData("24pt", 24)]
        [InlineData("16px", 12)]
        [InlineData("32px", 24)]
        [InlineData("1em", 12)]
        [InlineData("2em", 24)]
        [InlineData("1.5rem", 18)]
        [InlineData("200%", 24)]
        [InlineData("50%", 6)]
        [InlineData("10", 10)]
        public void ParseFontSize_VariousUnits(string input, int expected)
        {
            var method = GetStatic("ParseFontSize");
            var result = (int)method.Invoke(null, new object[] { input, 12 })!;
            Assert.Equal(expected, result);
        }

        [Theory]
        [InlineData("12pt", 240)]
        [InlineData("1em", 240)]
        [InlineData("2px", 20)]  // 2px -> 1pt -> 20 twips
        [InlineData("50%", 120)]
        public void ParseLengthToTwips_ConvertsCorrectly(string input, int expectedTwips)
        {
            var method = GetStatic("ParseLengthToTwips");
            var result = (int)method.Invoke(null, new object[] { input, 12 })!;
            Assert.Equal(expectedTwips, result);
        }

        [Theory]
        [InlineData("#ff0000", "#ff0000")]
        [InlineData("#f00", "#ff0000")]
        [InlineData("#ff000080", "#ff0000")]
        [InlineData("rgb(255,0,0)", "#ff0000")]
        [InlineData("rgba(255,0,0,0.5)", "#ff0000")]
        [InlineData("hsl(0,100%,50%)", "#ff0000")]
        [InlineData("hsla(240,100%,50%,0.3)", "#0000ff")]
        [InlineData("transparent", null)]
        [InlineData("red", "#ff0000")]
        [InlineData("blue", "#0000ff")]
        [InlineData("orange", "#ffa500")]
        [InlineData("unknown", null)]
        public void ParseColor_VariousFormats(string input, string? expected)
        {
            var method = GetStatic("ParseColor");
            var result = (string?)method.Invoke(null, new object[] { input });
            Assert.Equal(expected, result);
        }

        [Fact]
        public void HtmlParser_VoidBrTag_IsSelfClosing()
        {
            string html = "<p>a<br>b</p>";
            var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);

            string xml;
            using (var ms = new System.IO.MemoryStream(bytes))
            using (var z = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, true))
            {
                var entry = z.GetEntry("word/document.xml");
                using var r = new System.IO.StreamReader(entry!.Open());
                xml = r.ReadToEnd();
            }

            Assert.Contains("a", xml);
            Assert.Contains("<w:br/>", xml);
            Assert.Contains("b", xml);
        }

        [Fact]
        public void HtmlParser_WhitespaceBetweenInlineNodes_IsPreservedAsSpace()
        {
            string html = "<p><span>foo</span> <span>bar</span></p>";
            var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);

            string xml;
            using (var ms = new System.IO.MemoryStream(bytes))
            using (var z = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, true))
            {
                var entry = z.GetEntry("word/document.xml");
                using var r = new System.IO.StreamReader(entry!.Open());
                xml = r.ReadToEnd();
            }

            Assert.Contains("foo", xml);
            Assert.Contains("> </w:t>", xml);
            Assert.Contains("bar", xml);
        }

        [Fact]
        public void HtmlDocument_ImplicitLiClose_AllowsMissingEndTags()
        {
            // Real HTML often omits </li>
            string html = "<ul><li>a<li>b</ul>";
            var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);

            string xml;
            using (var ms = new System.IO.MemoryStream(bytes))
            using (var z = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, true))
            {
                var entry = z.GetEntry("word/document.xml");
                using var r = new System.IO.StreamReader(entry!.Open());
                xml = r.ReadToEnd();
            }

            Assert.Contains("a", xml);
            Assert.Contains("b", xml);
        }

        [Fact]
        public void HtmlDocument_MismatchedEndTags_DoNotDropText()
        {
            // End tag closes parent out of order; we should still keep text nodes.
            string html = "<div><p><span>x</div>y";
            var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);

            string xml;
            using (var ms = new System.IO.MemoryStream(bytes))
            using (var z = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, true))
            {
                var entry = z.GetEntry("word/document.xml");
                using var r = new System.IO.StreamReader(entry!.Open());
                xml = r.ReadToEnd();
            }

            Assert.Contains("x", xml);
            Assert.Contains("y", xml);
        }

        [Fact]
        public void HtmlDocument_TableMissingCellEndTags_StillParsesAllCells()
        {
            // Real HTML often omits </td>
            string html = "<table><tr><td>a<td>b</tr></table>";
            var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);

            string xml;
            using (var ms = new System.IO.MemoryStream(bytes))
            using (var z = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, true))
            {
                var entry = z.GetEntry("word/document.xml");
                using var r = new System.IO.StreamReader(entry!.Open());
                xml = r.ReadToEnd();
            }

            Assert.Contains("a", xml);
            Assert.Contains("b", xml);
        }

        [Fact]
        public void HtmlDocument_TableMissingRowEndTags_StillParsesAllRows()
        {
            // Missing </td>, </tr>, </tbody>
            string html = "<table><tbody><tr><td>a<tr><td>b</tbody></table>";
            var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);

            string xml;
            using (var ms = new System.IO.MemoryStream(bytes))
            using (var z = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, true))
            {
                var entry = z.GetEntry("word/document.xml");
                using var r = new System.IO.StreamReader(entry!.Open());
                xml = r.ReadToEnd();
            }

            Assert.Contains("a", xml);
            Assert.Contains("b", xml);
        }

        [Fact]
        public void NumberingXml_DefinesMultipleLevels_ForNestedLists()
        {
            string html = "<ul><li>a<ul><li>b<ul><li>c</li></ul></li></ul></li></ul>";
            var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);

            string numbering;
            using (var ms = new System.IO.MemoryStream(bytes))
            using (var z = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, true))
            {
                var entry = z.GetEntry("word/numbering.xml");
                using var r = new System.IO.StreamReader(entry!.Open());
                numbering = r.ReadToEnd();
            }

            Assert.Contains("w:ilvl=\"0\"", numbering);
            Assert.Contains("w:ilvl=\"1\"", numbering);
            Assert.Contains("w:ilvl=\"2\"", numbering);
            Assert.Contains("w:numFmt w:val=\"lowerLetter\"", numbering);
            Assert.Contains("w:numFmt w:val=\"lowerRoman\"", numbering);
            Assert.Contains("w:suff w:val=\"space\"", numbering);
        }

        [Fact]
        public void ListItem_TextAfterNestedList_IsNotLost()
        {
            // Ensures we resume the outer list item paragraph after a nested list block.
            string html = "<ul><li>one<ul><li>sub</li></ul>tail</li></ul>";
            var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);

            string xml;
            using (var ms = new System.IO.MemoryStream(bytes))
            using (var z = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, true))
            {
                var entry = z.GetEntry("word/document.xml");
                using var r = new System.IO.StreamReader(entry!.Open());
                xml = r.ReadToEnd();
            }

            Assert.Contains("one", xml);
            Assert.Contains("sub", xml);
            Assert.Contains("tail", xml);

            // Outer and inner list item markers should both exist, but "tail" should be a continuation paragraph
            // (i.e. not create a third list marker).
            int numPrCount = System.Text.RegularExpressions.Regex.Matches(xml, "<w:numPr>").Count;
            Assert.Equal(2, numPrCount);
        }

        [Fact]
        public void ListItem_MultipleParagraphs_DoNotCreateMultipleMarkers()
        {
            string html = "<ul><li><p>a</p><p>b</p></li></ul>";
            var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);

            string xml;
            using (var ms = new System.IO.MemoryStream(bytes))
            using (var z = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, true))
            {
                var entry = z.GetEntry("word/document.xml");
                using var r = new System.IO.StreamReader(entry!.Open());
                xml = r.ReadToEnd();
            }

            Assert.Contains("a", xml);
            Assert.Contains("b", xml);
            int numPrCount = System.Text.RegularExpressions.Regex.Matches(xml, "<w:numPr>").Count;
            Assert.Equal(1, numPrCount);

            // Continuation paragraph should align like list text: left=720, hanging=360 (so text start = 360).
            Assert.Contains("w:left=\"720\"", xml);
            Assert.Contains("w:hanging=\"360\"", xml);
        }

        [Fact]
        public void ListContinuationIndent_MatchesNumberingLevelLeftIndent()
        {
            string html = "<ul><li>a<ul><li>b</li></ul>c</li></ul>";
            var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);

            string docXml;
            string numberingXml;
            using (var ms = new System.IO.MemoryStream(bytes))
            using (var z = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, true))
            {
                using (var r = new System.IO.StreamReader(z.GetEntry("word/document.xml")!.Open()))
                    docXml = r.ReadToEnd();
                using (var r = new System.IO.StreamReader(z.GetEntry("word/numbering.xml")!.Open()))
                    numberingXml = r.ReadToEnd();
            }

            // level 0 left indent is 720, continuation should use the same.
            Assert.Contains("<w:lvl w:ilvl=\"0\">", numberingXml);
            Assert.Contains("w:left=\"720\"", numberingXml);
            Assert.Contains("w:left=\"720\"", docXml);
            Assert.Contains("w:hanging=\"360\"", docXml);
            Assert.Contains("<w:tab w:val=\"num\" w:pos=\"720\"/>", docXml);
            Assert.Contains("<w:pStyle w:val=\"ListParagraph\"/>", docXml);
        }

        [Fact]
        public void StylesXml_DefinesListParagraphStyle()
        {
            string html = "<ul><li>a</li></ul>";
            var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html);

            string styles;
            using (var ms = new System.IO.MemoryStream(bytes))
            using (var z = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, true))
            {
                var entry = z.GetEntry("word/styles.xml");
                using var r = new System.IO.StreamReader(entry!.Open());
                styles = r.ReadToEnd();
            }

            Assert.Contains("w:styleId=\"ListParagraph\"", styles);
        }
    }
}
