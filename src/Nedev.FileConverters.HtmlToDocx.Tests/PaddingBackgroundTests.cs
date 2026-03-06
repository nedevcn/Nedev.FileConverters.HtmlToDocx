using Xunit;
using Nedev.FileConverters.HtmlToDocx.Core.Conversion;
using System.IO.Compression;

namespace Nedev.FileConverters.HtmlToDocx.Tests
{
    public class PaddingBackgroundTests
    {
        [Fact]
        public void ParagraphPaddingAndBackgroundApplied()
        {
            string html = "<p style=\"padding-left:5px;padding-right:10px;background-color:#0f0\">Hi</p>";
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
            // padding-left 5px => 5px -> 3pt (rounded) -> 60 twips
            // padding-right 10px => 7.5pt -> 7pt -> 140 twips
            Assert.Contains("w:ind", xml);
            Assert.Contains("w:left=\"60\"", xml);
            Assert.Contains("w:right=\"140\"", xml);
            Assert.Contains("w:shd", xml);
            Assert.Contains("fill=\"00ff00\"", xml);
        }
    }
}