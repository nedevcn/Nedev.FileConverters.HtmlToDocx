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
        public void MatchesSelector_AttributeSelectors()
        {
            var node = MakeNode("p");
            node.Attributes["data-foo"] = "bar";
            Assert.True(InvokeMatches(node, "[data-foo]") );
            Assert.True(InvokeMatches(node, "p[data-foo=bar]") );
            Assert.False(InvokeMatches(node, "[data-foo=baz]") );
        }

        [Fact]
        public void MatchesSelector_SiblingCombinators()
        {
            var parent = MakeNode("div");
            var a = MakeNode("p");
            var b = MakeNode("span");
            var c = MakeNode("p");
            a.Parent = parent;
            b.Parent = parent;
            c.Parent = parent;
            parent.Children.Add(a);
            parent.Children.Add(b);
            parent.Children.Add(c);

            Assert.True(InvokeMatches(b, "p + span"));
            Assert.False(InvokeMatches(c, "p + span"));

            Assert.True(InvokeMatches(c, "p ~ p"));
            Assert.False(InvokeMatches(a, "p ~ p"));
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

        [Fact]
        public void ParseDeclarations_ReadsImportantFlag()
        {
            var declarations = CssParser.ParseDeclarations("color: red !important; font-size: 12pt;");
            Assert.Equal(2, declarations.Count);
            Assert.True(declarations[0].Important);
            Assert.Equal("red", declarations[0].Value);
            Assert.False(declarations[1].Important);
        }

        [Fact]
        public void ResolveStyles_UsesSpecificity()
        {
            var doc = HtmlDocument.Parse("<style>p{color:blue;} #x{color:red;}</style><p id=\"x\">t</p>");
            StyleResolver.ResolveStyles(doc);
            var node = FindFirstElementByTag(doc.Root, "p");
            Assert.NotNull(node);
            Assert.True(node!.ComputedStyle.TryGetValue("color", out var color));
            Assert.Equal("red", color);
        }

        [Fact]
        public void ResolveStyles_ImportantBeatsInlineNonImportant()
        {
            var doc = HtmlDocument.Parse("<style>#x{color:red;} p{color:blue !important;}</style><p id=\"x\" style=\"color:green\">t</p>");
            StyleResolver.ResolveStyles(doc);
            var node = FindFirstElementByTag(doc.Root, "p");
            Assert.NotNull(node);
            Assert.True(node!.ComputedStyle.TryGetValue("color", out var color));
            Assert.Equal("blue", color);
        }

        [Fact]
        public void ResolveStyles_InlineImportantBeatsStylesheetImportant()
        {
            var doc = HtmlDocument.Parse("<style>#x{color:red !important;}</style><p id=\"x\" style=\"color:green !important\">t</p>");
            StyleResolver.ResolveStyles(doc);
            var node = FindFirstElementByTag(doc.Root, "p");
            Assert.NotNull(node);
            Assert.True(node!.ComputedStyle.TryGetValue("color", out var color));
            Assert.Equal("green", color);
        }

        private bool InvokeMatches(HtmlNode node, string selector)
        {
            // Use reflection to call the internal method
            var method = typeof(StyleResolver).GetMethod("MatchesSelector", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            return (bool)method!.Invoke(null, new object[] { node, selector });
        }

        private static HtmlNode? FindFirstElementByTag(HtmlNode node, string tagName)
        {
            if (node.TagName.Equals(tagName, StringComparison.OrdinalIgnoreCase))
                return node;

            foreach (var child in node.Children)
            {
                var found = FindFirstElementByTag(child, tagName);
                if (found != null) return found;
            }

            return null;
        }
    }
}
