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
    }
}