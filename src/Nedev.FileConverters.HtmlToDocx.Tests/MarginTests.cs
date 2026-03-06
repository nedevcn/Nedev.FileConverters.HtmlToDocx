using Xunit;
using Nedev.FileConverters.HtmlToDocx.Core.Html;
using Nedev.FileConverters.HtmlToDocx.Core.Conversion;

namespace Nedev.FileConverters.HtmlToDocx.Tests
{
    public class MarginTests
    {
        [Fact]
        public void ParagraphMarginIsConvertedToSpacing()
        {
            string html = "<p style=\"margin:10px 0 20px 0\">Text</p>";
            var doc = HtmlDocument.Parse(html);
            var converter = new HtmlToDocxConverter();
            var bytes = converter.Convert(html); // we just need resulting xml from builder

            // convert bytes to string and inspect
            string xml;
            using (var ms = new System.IO.MemoryStream(bytes))
            using (var z = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read, true))
            {
                var entry = z.GetEntry("word/document.xml");
                using var r = new System.IO.StreamReader(entry.Open());
                xml = r.ReadToEnd();
            }

            // compute expected twips
            var parseLen = typeof(HtmlToDocxConverter).GetMethod("ParseLengthToTwips", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            var before = (int)parseLen!.Invoke(null, new object[] { "10px", 12 })!;
            var after = (int)parseLen!.Invoke(null, new object[] { "20px", 12 })!;
            Assert.Contains($"w:before=\"{before}\"", xml);
            Assert.Contains($"w:after=\"{after}\"", xml);
        }
    }
}