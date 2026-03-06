using Xunit;
using Nedev.FileConverters.HtmlToDocx.Core.Conversion;
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
    }
}