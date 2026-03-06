using Xunit;
using Nedev.FileConverters.HtmlToDocx.Core.Css;
using Nedev.FileConverters.HtmlToDocx.Core.Html;

namespace Nedev.FileConverters.HtmlToDocx.Tests
{
    public class CssResolverTests
    {
        private HtmlNode MakeNode(string tag, string? id = null, string? @class = null)
        {
            var node = new HtmlNode { TagName = tag };
            if (id != null) node.Attributes["id"] = id;
            if (@class != null) node.Attributes["class"] = @class;
            return node;
        }

        [Fact]
        public void ParseStylesheet_SplitsCommaSelectors()
        {
            var css = "h1, h2 { color: red; }";
            var rules = CssParser.ParseStylesheet(css);
            Assert.Equal(2, rules.Count);
            Assert.Contains(rules, r => r.Selector == "h1");
            Assert.Contains(rules, r => r.Selector == "h2");
        }

        [Fact]
        public void MatchesSelector_TagAndClassAndId()
        {
            var node = MakeNode("div", "myid", "foo bar");
            Assert.True(InvokeMatches(node, "div"));
            Assert.True(InvokeMatches(node, "#myid"));
            Assert.True(InvokeMatches(node, ".foo"));
            Assert.True(InvokeMatches(node, "div.foo"));
            Assert.True(InvokeMatches(node, "div#myid"));
            Assert.False(InvokeMatches(node, "span"));
            Assert.False(InvokeMatches(node, ".baz"));
        }

        [Fact]
        public void MatchesSelector_DescendantCombinator()
        {
            var root = MakeNode("div");
            var child = MakeNode("span");
            child.Parent = root;
            root.Children.Add(child);

            Assert.True(InvokeMatches(child, "div span"));
            Assert.False(InvokeMatches(root, "div span"));
        }

        [Fact]
        public void MatchesSelector_ChildCombinator()
        {
            var root = MakeNode("div");
            var child = MakeNode("p");
            child.Parent = root;
            root.Children.Add(child);
            var grand = MakeNode("span");
            grand.Parent = child;
            child.Children.Add(grand);

            // diagnostic information - inspect tokens and internal recursion
            var matcher = typeof(StyleResolver).GetMethod("MatchesSelector", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            var recurse = typeof(StyleResolver).GetMethod("MatchSelectorParts", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            string sel = "div > p";
            var tokens = sel.Split(' ', System.StringSplitOptions.RemoveEmptyEntries);
            System.Console.WriteLine("tokens: " + string.Join("|", tokens));
            bool resultChild = (bool)matcher!.Invoke(null, new object[] { child, sel });
            Assert.True(resultChild, $"selector '{sel}' on child returned {resultChild}");
            // also try calling the recursive helper directly for insight
            bool recursiveResult = (bool)recurse!.Invoke(null, new object?[] { child, tokens, tokens.Length - 1 });
            System.Console.WriteLine($"recursive result: {recursiveResult}");

            Assert.False(InvokeMatches(grand, "div > p"));
            // grand is a span under p, so descendant rule uses span
            Assert.True(InvokeMatches(grand, "div span"));
        }

        private bool InvokeMatches(HtmlNode node, string selector)
        {
            // Use reflection to call the internal method
            var method = typeof(StyleResolver).GetMethod("MatchesSelector", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            return (bool)method!.Invoke(null, new object[] { node, selector });
        }
    }
}