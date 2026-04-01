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
        public void MatchesSelector_AttributeOperators()
        {
            var node = MakeNode("p");
            node.Attributes["data-id"] = "pre-mid-suf";
            node.Attributes["class"] = "alpha beta";
            node.Attributes["lang"] = "en-US";

            Assert.True(InvokeMatches(node, "[data-id^=pre]"));
            Assert.True(InvokeMatches(node, "[data-id$=suf]"));
            Assert.True(InvokeMatches(node, "[data-id*=mid]"));
            Assert.True(InvokeMatches(node, "[class~=beta]"));
            Assert.True(InvokeMatches(node, "[lang|=en]"));
            Assert.False(InvokeMatches(node, "[data-id^=x]"));
            Assert.False(InvokeMatches(node, "[class~=gamma]"));
        }

        [Fact]
        public void MatchesSelector_NotPseudoClass()
        {
            var node = MakeNode("div", @class: "foo bar");
            Assert.True(InvokeMatches(node, "div:not(.baz)"));
            Assert.False(InvokeMatches(node, "div:not(.foo)"));
            Assert.True(InvokeMatches(node, "div:not(#x)"));
            Assert.False(InvokeMatches(node, "div:not(div.foo)"));
        }

        [Fact]
        public void MatchesSelector_IsAndWherePseudoClass()
        {
            var node = MakeNode("div", "id1", "foo bar");
            Assert.True(InvokeMatches(node, "div:is(.foo,#x)"));
            Assert.True(InvokeMatches(node, "div:where(.foo,#x)"));
            Assert.False(InvokeMatches(node, "div:is(.baz,#x)"));
            Assert.False(InvokeMatches(node, "div:where(.baz,#x)"));
        }

        [Fact]
        public void MatchesSelector_StructuralPseudoClasses()
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

            Assert.True(InvokeMatches(a, "p:first-child"));
            Assert.True(InvokeMatches(c, "p:last-child"));
            Assert.True(InvokeMatches(b, "span:only-of-type"));
            Assert.False(InvokeMatches(a, "p:only-child"));
            Assert.True(InvokeMatches(c, ":nth-child(3)"));
            Assert.True(InvokeMatches(c, "p:nth-of-type(2)"));
            Assert.True(InvokeMatches(a, ":nth-child(odd)"));
            Assert.True(InvokeMatches(c, ":nth-child(2n+1)"));
        }

        [Fact]
        public void MatchesSelector_NthLastAndEmptyPseudoClasses()
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

            Assert.True(InvokeMatches(a, ":nth-last-child(3)"));
            Assert.True(InvokeMatches(b, ":nth-last-child(2)"));
            Assert.True(InvokeMatches(a, "p:nth-last-of-type(2)"));
            Assert.True(InvokeMatches(c, "p:nth-last-of-type(1)"));

            var empty = MakeNode("div");
            var nonEmpty = MakeNode("div");
            var txt = new HtmlNode { TagName = "#text", Text = "x", Parent = nonEmpty };
            nonEmpty.Children.Add(txt);
            Assert.True(InvokeMatches(empty, "div:empty"));
            Assert.False(InvokeMatches(nonEmpty, "div:empty"));
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

        [Fact]
        public void ResolveStyles_NotPseudoSpecificityAndMatch()
        {
            var doc = HtmlDocument.Parse("<style>p{color:black;} p:not(.x){color:blue;} #id{color:red;}</style><p id=\"id\" class=\"x\">t</p><p>u</p>");
            StyleResolver.ResolveStyles(doc);
            var allP = GetElementsByTag(doc.Root, "p");
            Assert.Equal(2, allP.Count);
            Assert.True(allP[0].ComputedStyle.TryGetValue("color", out var first));
            Assert.True(allP[1].ComputedStyle.TryGetValue("color", out var second));
            Assert.Equal("red", first);
            Assert.Equal("blue", second);
        }

        [Fact]
        public void ResolveStyles_IsPseudoUsesMaxSpecificityOfArguments()
        {
            var doc = HtmlDocument.Parse("<style>p{color:black;} #id{color:red;} p:is(.a,#id){color:blue;}</style><p id=\"id\" class=\"a\">t</p>");
            StyleResolver.ResolveStyles(doc);
            var node = FindFirstElementByTag(doc.Root, "p");
            Assert.NotNull(node);
            Assert.True(node!.ComputedStyle.TryGetValue("color", out var color));
            Assert.Equal("blue", color);
        }

        [Fact]
        public void ResolveStyles_WherePseudoHasZeroSpecificity()
        {
            var doc = HtmlDocument.Parse("<style>p{color:black;} p:where(.a,#id){color:blue;} #id{color:red;}</style><p id=\"id\" class=\"a\">t</p><p class=\"a\">u</p>");
            StyleResolver.ResolveStyles(doc);
            var allP = GetElementsByTag(doc.Root, "p");
            Assert.Equal(2, allP.Count);
            Assert.True(allP[0].ComputedStyle.TryGetValue("color", out var first));
            Assert.True(allP[1].ComputedStyle.TryGetValue("color", out var second));
            Assert.Equal("red", first);
            Assert.Equal("blue", second);
        }

        [Fact]
        public void ResolveStyles_NthChildAndNthOfType()
        {
            var doc = HtmlDocument.Parse("<style>li{color:black;} li:nth-child(2){color:red;} li:nth-of-type(3){color:blue;}</style><ul><li>a</li><li>b</li><li>c</li></ul>");
            StyleResolver.ResolveStyles(doc);
            var allLi = GetElementsByTag(doc.Root, "li");
            Assert.Equal(3, allLi.Count);
            Assert.True(allLi[0].ComputedStyle.TryGetValue("color", out var c1));
            Assert.True(allLi[1].ComputedStyle.TryGetValue("color", out var c2));
            Assert.True(allLi[2].ComputedStyle.TryGetValue("color", out var c3));
            Assert.Equal("black", c1);
            Assert.Equal("red", c2);
            Assert.Equal("blue", c3);
        }

        [Fact]
        public void ResolveStyles_NthLastAndEmpty()
        {
            var doc = HtmlDocument.Parse("<style>li{color:black;} li:nth-last-child(1){color:blue;} li:nth-last-of-type(2){background-color:yellow;} div:empty{color:green;}</style><ul><li>a</li><li>b</li><li>c</li></ul><div></div><div>x</div>");
            StyleResolver.ResolveStyles(doc);
            var allLi = GetElementsByTag(doc.Root, "li");
            Assert.Equal(3, allLi.Count);
            Assert.Equal("blue", allLi[2].ComputedStyle["color"]);
            Assert.Equal("yellow", allLi[1].ComputedStyle["background-color"]);

            var allDiv = GetElementsByTag(doc.Root, "div");
            Assert.True(allDiv[0].ComputedStyle.TryGetValue("color", out var c0));
            Assert.Equal("green", c0);
            Assert.False(allDiv[1].ComputedStyle.ContainsKey("color"));
        }

        [Fact]
        public void MatchesSelector_HasPseudoClass()
        {
            var root = MakeNode("div", "root");
            var child = MakeNode("section");
            var leaf = MakeNode("span", @class: "target");
            child.Parent = root;
            leaf.Parent = child;
            root.Children.Add(child);
            child.Children.Add(leaf);

            Assert.True(InvokeMatches(root, "div:has(.target)"));
            Assert.True(InvokeMatches(root, "div:has(> section)"));
            Assert.True(InvokeMatches(root, "div:has(section .target)"));
            Assert.False(InvokeMatches(child, "section:has(> p)"));
        }

        [Fact]
        public void MatchesSelector_HasLeadingCombinatorBoundary()
        {
            var root = MakeNode("div");
            var child = MakeNode("section");
            var nested = MakeNode("section");
            child.Parent = root;
            nested.Parent = child;
            root.Children.Add(child);
            child.Children.Add(nested);

            Assert.True(InvokeMatches(root, "div:has(> section)"));
            Assert.False(InvokeMatches(root, "div:has(> article)"));
            Assert.False(InvokeMatches(root, "div:has(> article, > nav)"));
            Assert.False(InvokeMatches(root, "div:has(> section > article)"));
        }

        [Fact]
        public void ResolveStyles_HasPseudoSpecificity()
        {
            var doc = HtmlDocument.Parse("<style>div{color:black;} #x{color:red;} div:has(#in){color:blue;}</style><div id=\"x\"><span id=\"in\"></span></div>");
            StyleResolver.ResolveStyles(doc);
            var node = FindFirstElementByTag(doc.Root, "div");
            Assert.NotNull(node);
            Assert.True(node!.ComputedStyle.TryGetValue("color", out var color));
            Assert.Equal("blue", color);
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

        private static System.Collections.Generic.List<HtmlNode> GetElementsByTag(HtmlNode node, string tagName)
        {
            var result = new System.Collections.Generic.List<HtmlNode>();
            void Walk(HtmlNode n)
            {
                if (n.TagName.Equals(tagName, StringComparison.OrdinalIgnoreCase))
                    result.Add(n);
                foreach (var c in n.Children) Walk(c);
            }
            Walk(node);
            return result;
        }
    }
}
