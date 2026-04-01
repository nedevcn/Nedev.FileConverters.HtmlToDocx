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
        public void MatchesSelector_NotPseudoClass_ComplexSelectorArgument()
        {
            var parent = MakeNode("div");
            var child = MakeNode("span");
            child.Parent = parent;
            parent.Children.Add(child);

            Assert.False(InvokeMatches(child, "span:not(div > span)"));
            Assert.True(InvokeMatches(child, "span:not(p > span)"));
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
        public void ResolveStyles_NotPseudoWithComplexSelectorArgument()
        {
            var doc = HtmlDocument.Parse("<style>span{color:black;} span:not(div > span){color:blue;}</style><div><span id=\"a\">a</span></div><p><span id=\"b\">b</span></p>");
            StyleResolver.ResolveStyles(doc);
            var a = FindFirstElementById(doc.Root, "a");
            var b = FindFirstElementById(doc.Root, "b");
            Assert.NotNull(a);
            Assert.NotNull(b);
            Assert.Equal("black", a!.ComputedStyle["color"]);
            Assert.Equal("blue", b!.ComputedStyle["color"]);
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
            var article = MakeNode("article");
            child.Parent = root;
            nested.Parent = child;
            article.Parent = child;
            root.Children.Add(child);
            child.Children.Add(nested);
            child.Children.Add(article);

            Assert.True(InvokeMatches(root, "div:has(> section)"));
            Assert.False(InvokeMatches(root, "div:has(> article)"));
            Assert.False(InvokeMatches(root, "div:has(> article, > nav)"));
            Assert.True(InvokeMatches(root, "div:has(> section > article)"));
            Assert.False(InvokeMatches(root, "div:has(> section > nav)"));
        }

        [Fact]
        public void MatchesSelector_RootAndScopePseudoClass()
        {
            var root = MakeNode("html");
            var body = MakeNode("body");
            var p = MakeNode("p");
            var document = new HtmlNode { TagName = "#document" };

            root.Parent = document;
            body.Parent = root;
            p.Parent = body;
            document.Children.Add(root);
            root.Children.Add(body);
            body.Children.Add(p);

            Assert.True(InvokeMatches(root, ":root"));
            Assert.True(InvokeMatches(root, "html:scope"));
            Assert.False(InvokeMatches(body, ":root"));
            Assert.False(InvokeMatches(p, ":scope"));
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

        [Fact]
        public void ResolveStyles_RootAndScopePseudoApplyToDocumentRootElement()
        {
            var doc = HtmlDocument.Parse("<style>:root{color:red;} :scope{background-color:yellow;} body{color:blue;}</style><html><body><p>t</p></body></html>");
            StyleResolver.ResolveStyles(doc);
            var html = FindFirstElementByTag(doc.Root, "html");
            var body = FindFirstElementByTag(doc.Root, "body");
            Assert.NotNull(html);
            Assert.NotNull(body);
            Assert.Equal("red", html!.ComputedStyle["color"]);
            Assert.Equal("yellow", html.ComputedStyle["background-color"]);
            Assert.Equal("blue", body!.ComputedStyle["color"]);
        }

        [Fact]
        public void ResolveStyles_SpecificityComponentComparison_ClassOverflowDoesNotBeatId()
        {
            var doc = HtmlDocument.Parse("<style>#x{color:red;}.c1.c2.c3.c4.c5.c6.c7.c8.c9.c10{color:blue;}</style><p id=\"x\" class=\"c1 c2 c3 c4 c5 c6 c7 c8 c9 c10\">t</p>");
            StyleResolver.ResolveStyles(doc);
            var node = FindFirstElementByTag(doc.Root, "p");
            Assert.NotNull(node);
            Assert.True(node!.ComputedStyle.TryGetValue("color", out var color));
            Assert.Equal("red", color);
        }

        [Fact]
        public void MatchesSelector_LangAndDirPseudoClass()
        {
            var html = MakeNode("html");
            html.Attributes["lang"] = "en-US";
            html.Attributes["dir"] = "rtl";
            var body = MakeNode("body");
            var p = MakeNode("p");
            body.Parent = html;
            p.Parent = body;
            html.Children.Add(body);
            body.Children.Add(p);

            Assert.True(InvokeMatches(p, ":lang(en)"));
            Assert.True(InvokeMatches(p, ":lang(en-US)"));
            Assert.False(InvokeMatches(p, ":lang(fr)"));
            Assert.True(InvokeMatches(p, ":dir(rtl)"));
            Assert.False(InvokeMatches(p, ":dir(ltr)"));
        }

        [Fact]
        public void ResolveStyles_LangAndDirPseudoClassApplyByInheritance()
        {
            var doc = HtmlDocument.Parse("<style>:lang(en){color:red;} :dir(rtl){text-align:right;} p{color:black;}</style><div lang=\"en-GB\" dir=\"rtl\"><p>t</p></div>");
            StyleResolver.ResolveStyles(doc);
            var p = FindFirstElementByTag(doc.Root, "p");
            Assert.NotNull(p);
            Assert.Equal("red", p!.ComputedStyle["color"]);
            Assert.Equal("right", p.ComputedStyle["text-align"]);
        }

        [Fact]
        public void MatchesSelector_FormStatePseudoClass()
        {
            var form = MakeNode("form");
            var inputDisabled = MakeNode("input");
            inputDisabled.Attributes["disabled"] = "disabled";
            var inputEnabled = MakeNode("input");
            var inputChecked = MakeNode("input");
            inputChecked.Attributes["checked"] = "checked";
            var optionSelected = MakeNode("option");
            optionSelected.Attributes["selected"] = "selected";

            inputDisabled.Parent = form;
            inputEnabled.Parent = form;
            inputChecked.Parent = form;
            optionSelected.Parent = form;
            form.Children.Add(inputDisabled);
            form.Children.Add(inputEnabled);
            form.Children.Add(inputChecked);
            form.Children.Add(optionSelected);

            Assert.True(InvokeMatches(inputDisabled, "input:disabled"));
            Assert.True(InvokeMatches(inputEnabled, "input:enabled"));
            Assert.True(InvokeMatches(inputChecked, "input:checked"));
            Assert.True(InvokeMatches(optionSelected, "option:checked"));
            Assert.False(InvokeMatches(inputEnabled, "input:checked"));
        }

        [Fact]
        public void ResolveStyles_FormStatePseudoClass()
        {
            var doc = HtmlDocument.Parse("<style>input:disabled{color:red;} input:enabled{color:blue;} input:checked{font-weight:bold;} option:checked{text-decoration:underline;}</style><div><input disabled /><input checked /><input /><select><option selected>v</option></select></div>");
            StyleResolver.ResolveStyles(doc);

            var inputs = GetElementsByTag(doc.Root, "input");
            var options = GetElementsByTag(doc.Root, "option");
            Assert.Equal(3, inputs.Count);
            Assert.Single(options);

            Assert.Equal("red", inputs[0].ComputedStyle["color"]);
            Assert.Equal("blue", inputs[1].ComputedStyle["color"]);
            Assert.Equal("bold", inputs[1].ComputedStyle["font-weight"]);
            Assert.Equal("blue", inputs[2].ComputedStyle["color"]);
            Assert.Equal("underline", options[0].ComputedStyle["text-decoration"]);
        }

        [Fact]
        public void MatchesSelector_FormConstraintAndEditabilityPseudoClass()
        {
            var required = MakeNode("input");
            required.Attributes["required"] = "required";
            var optional = MakeNode("input");
            var readonlyInput = MakeNode("input");
            readonlyInput.Attributes["readonly"] = "readonly";
            var textarea = MakeNode("textarea");
            var editableDiv = MakeNode("div");
            editableDiv.Attributes["contenteditable"] = "true";

            Assert.True(InvokeMatches(required, "input:required"));
            Assert.True(InvokeMatches(optional, "input:optional"));
            Assert.True(InvokeMatches(readonlyInput, "input:read-only"));
            Assert.True(InvokeMatches(textarea, "textarea:read-write"));
            Assert.True(InvokeMatches(editableDiv, "div:read-write"));
            Assert.False(InvokeMatches(optional, "input:required"));
            Assert.False(InvokeMatches(required, "input:optional"));
        }

        [Fact]
        public void ResolveStyles_FormConstraintAndEditabilityPseudoClass()
        {
            var doc = HtmlDocument.Parse("<style>input:required{color:red;} input:optional{color:blue;} input:read-only{font-style:italic;} textarea:read-write{text-decoration:underline;} div:read-write{font-weight:bold;}</style><div><input required /><input readonly /><input /><textarea></textarea><div contenteditable=\"true\">x</div></div>");
            StyleResolver.ResolveStyles(doc);

            var inputs = GetElementsByTag(doc.Root, "input");
            var textarea = GetElementsByTag(doc.Root, "textarea").Single();
            var divs = GetElementsByTag(doc.Root, "div");
            var editableDiv = divs.Last();

            Assert.Equal("red", inputs[0].ComputedStyle["color"]);
            Assert.Equal("italic", inputs[1].ComputedStyle["font-style"]);
            Assert.Equal("blue", inputs[2].ComputedStyle["color"]);
            Assert.Equal("underline", textarea.ComputedStyle["text-decoration"]);
            Assert.Equal("bold", editableDiv.ComputedStyle["font-weight"]);
        }

        [Fact]
        public void MatchesSelector_DisabledInheritance_OptgroupAndFieldsetLegendException()
        {
            var fieldset = MakeNode("fieldset");
            fieldset.Attributes["disabled"] = "disabled";
            var legend = MakeNode("legend");
            var inputInLegend = MakeNode("input");
            var inputInFieldset = MakeNode("input");
            legend.Parent = fieldset;
            inputInLegend.Parent = legend;
            inputInFieldset.Parent = fieldset;
            fieldset.Children.Add(legend);
            legend.Children.Add(inputInLegend);
            fieldset.Children.Add(inputInFieldset);

            var select = MakeNode("select");
            var optgroup = MakeNode("optgroup");
            optgroup.Attributes["disabled"] = "disabled";
            var option = MakeNode("option");
            optgroup.Parent = select;
            option.Parent = optgroup;
            select.Children.Add(optgroup);
            optgroup.Children.Add(option);

            Assert.False(InvokeMatches(inputInLegend, "input:disabled"));
            Assert.True(InvokeMatches(inputInLegend, "input:enabled"));
            Assert.True(InvokeMatches(inputInFieldset, "input:disabled"));
            Assert.True(InvokeMatches(option, "option:disabled"));
        }

        [Fact]
        public void ResolveStyles_DisabledInheritance_OptgroupAndFieldsetLegendException()
        {
            var doc = HtmlDocument.Parse("<style>input:disabled{color:red;} input:enabled{color:blue;} option:disabled{background-color:gray;} option:enabled{background-color:white;}</style><fieldset disabled><legend><input id=\"in-legend\" /></legend><input id=\"in-fieldset\" /></fieldset><select><optgroup disabled><option id=\"opt-a\">a</option></optgroup><option id=\"opt-b\">b</option></select><select disabled><option id=\"opt-c\">c</option></select>");
            StyleResolver.ResolveStyles(doc);

            var inLegend = FindFirstElementById(doc.Root, "in-legend");
            var inFieldset = FindFirstElementById(doc.Root, "in-fieldset");
            var optA = FindFirstElementById(doc.Root, "opt-a");
            var optB = FindFirstElementById(doc.Root, "opt-b");
            var optC = FindFirstElementById(doc.Root, "opt-c");

            Assert.NotNull(inLegend);
            Assert.NotNull(inFieldset);
            Assert.NotNull(optA);
            Assert.NotNull(optB);
            Assert.NotNull(optC);

            Assert.Equal("blue", inLegend!.ComputedStyle["color"]);
            Assert.Equal("red", inFieldset!.ComputedStyle["color"]);
            Assert.Equal("gray", optA!.ComputedStyle["background-color"]);
            Assert.Equal("white", optB!.ComputedStyle["background-color"]);
            Assert.Equal("gray", optC!.ComputedStyle["background-color"]);
        }

        [Fact]
        public void MatchesSelector_LinkAndAnyLinkPseudoClass()
        {
            var a = MakeNode("a");
            a.Attributes["href"] = "https://example.com";
            var area = MakeNode("area");
            area.Attributes["href"] = "/map";
            var link = MakeNode("link");
            link.Attributes["href"] = "/style.css";
            var noHref = MakeNode("a");

            Assert.True(InvokeMatches(a, "a:link"));
            Assert.True(InvokeMatches(a, "a:any-link"));
            Assert.True(InvokeMatches(area, "area:any-link"));
            Assert.True(InvokeMatches(link, "link:link"));
            Assert.False(InvokeMatches(noHref, "a:link"));
        }

        [Fact]
        public void ResolveStyles_LinkAndAnyLinkPseudoClass()
        {
            var doc = HtmlDocument.Parse("<style>a:link{color:red;} a:any-link{text-decoration:underline;} a{color:black;}</style><div><a href=\"https://example.com\">ok</a><a>no</a></div>");
            StyleResolver.ResolveStyles(doc);

            var links = GetElementsByTag(doc.Root, "a");
            Assert.Equal(2, links.Count);

            Assert.Equal("red", links[0].ComputedStyle["color"]);
            Assert.Equal("underline", links[0].ComputedStyle["text-decoration"]);
            Assert.Equal("black", links[1].ComputedStyle["color"]);
            Assert.False(links[1].ComputedStyle.ContainsKey("text-decoration"));
        }

        [Fact]
        public void ResolveStyles_CssVariables_InheritanceAndFallback()
        {
            var doc = HtmlDocument.Parse("<style>:root{--main:#123456;} div{color:var(--main);} p{color:var(--missing, green);}</style><html><body><div id=\"d\">x</div><p id=\"p\">y</p></body></html>");
            StyleResolver.ResolveStyles(doc);

            var d = FindFirstElementById(doc.Root, "d");
            var p = FindFirstElementById(doc.Root, "p");
            Assert.NotNull(d);
            Assert.NotNull(p);
            Assert.Equal("#123456", d!.ComputedStyle["color"]);
            Assert.Equal("green", p!.ComputedStyle["color"]);
        }

        [Fact]
        public void ResolveStyles_CssVariables_WorkWhenDefinedAfterUsageInSameRule()
        {
            var doc = HtmlDocument.Parse("<style>div{color:var(--x);--x:red;}</style><div id=\"x\">t</div>");
            StyleResolver.ResolveStyles(doc);
            var d = FindFirstElementById(doc.Root, "x");
            Assert.NotNull(d);
            Assert.Equal("red", d!.ComputedStyle["color"]);
        }

        [Fact]
        public void ResolveStyles_MediaPrintRules_Applied()
        {
            var doc = HtmlDocument.Parse("<style>p{color:black;} @media print { p { color:red; } }</style><p id=\"p\">t</p>");
            StyleResolver.ResolveStyles(doc);
            var p = FindFirstElementById(doc.Root, "p");
            Assert.NotNull(p);
            Assert.Equal("red", p!.ComputedStyle["color"]);
        }

        [Fact]
        public void ResolveStyles_MediaScreenRules_Ignored()
        {
            var doc = HtmlDocument.Parse("<style>@media screen { p { color:red; } } p{color:black;}</style><p id=\"p\">t</p>");
            StyleResolver.ResolveStyles(doc);
            var p = FindFirstElementById(doc.Root, "p");
            Assert.NotNull(p);
            Assert.Equal("black", p!.ComputedStyle["color"]);
        }

        [Fact]
        public void ResolveStyles_ParseDeclarations_DataUriSemicolonSafe()
        {
            var doc = HtmlDocument.Parse("<style>p{background-image:url(\"data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg'></svg>\");color:red;}</style><p id=\"p\">t</p>");
            StyleResolver.ResolveStyles(doc);
            var p = FindFirstElementById(doc.Root, "p");
            Assert.NotNull(p);
            Assert.Equal("red", p!.ComputedStyle["color"]);
            Assert.Contains("data:image/svg+xml;utf8", p.ComputedStyle["background-image"]);
        }

        [Fact]
        public void ResolveStyles_ParseDeclarations_QuotedSemicolonSafe()
        {
            var doc = HtmlDocument.Parse("<style>p{font-family:\"A;B\";color:blue;}</style><p id=\"p\">t</p>");
            StyleResolver.ResolveStyles(doc);
            var p = FindFirstElementById(doc.Root, "p");
            Assert.NotNull(p);
            Assert.Equal("blue", p!.ComputedStyle["color"]);
            Assert.Equal("\"A;B\"", p.ComputedStyle["font-family"]);
        }

        [Fact]
        public void ResolveStyles_CalcSameUnitExpression_Evaluated()
        {
            var doc = HtmlDocument.Parse("<style>p{width:calc(100px - 30px + 5px);}</style><p id=\"p\">t</p>");
            StyleResolver.ResolveStyles(doc);
            var p = FindFirstElementById(doc.Root, "p");
            Assert.NotNull(p);
            Assert.Equal("75px", p!.ComputedStyle["width"]);
        }

        [Fact]
        public void ResolveStyles_CalcWithVariables_EvaluatedAfterVarResolution()
        {
            var doc = HtmlDocument.Parse("<style>:root{--gap:12px;} p{margin-left:calc(var(--gap) + 8px);}</style><p id=\"p\">t</p>");
            StyleResolver.ResolveStyles(doc);
            var p = FindFirstElementById(doc.Root, "p");
            Assert.NotNull(p);
            Assert.Equal("20px", p!.ComputedStyle["margin-left"]);
        }

        [Fact]
        public void ResolveStyles_CalcMixedUnits_KeptAsOriginal()
        {
            var doc = HtmlDocument.Parse("<style>p{width:calc(100% - 20px);}</style><p id=\"p\">t</p>");
            StyleResolver.ResolveStyles(doc);
            var p = FindFirstElementById(doc.Root, "p");
            Assert.NotNull(p);
            Assert.Equal("calc(100% - 20px)", p!.ComputedStyle["width"]);
        }

        [Fact]
        public void ResolveStyles_SupportsRule_BasicAndNot()
        {
            var doc = HtmlDocument.Parse("<style>p{color:black;} @supports (color: red) { p { color: blue; } } @supports not (color: red) { p { color: green; } }</style><p id=\"p\">t</p>");
            StyleResolver.ResolveStyles(doc);
            var p = FindFirstElementById(doc.Root, "p");
            Assert.NotNull(p);
            Assert.Equal("blue", p!.ComputedStyle["color"]);
        }

        [Fact]
        public void ResolveStyles_SupportsRule_AndOr()
        {
            var doc = HtmlDocument.Parse("<style>p{color:black;} @supports ((color:red) and (font-size:12px)) { p { color: blue; } } @supports ((display:grid) or (display:flex)) { p { color: red; } }</style><p id=\"p\">t</p>");
            StyleResolver.ResolveStyles(doc);
            var p = FindFirstElementById(doc.Root, "p");
            Assert.NotNull(p);
            Assert.Equal("blue", p!.ComputedStyle["color"]);
        }

        [Fact]
        public void ResolveStyles_CssWideKeywords_InheritUnsetInitial()
        {
            var doc = HtmlDocument.Parse("<style>div{color:#112233;margin-top:12px;} p{color:inherit;margin-top:inherit;} span{color:unset;} em{margin-top:unset;} b{color:initial;}</style><div><p id=\"p\">a</p><span id=\"s\">b</span><em id=\"e\">c</em><b id=\"b\">d</b></div>");
            StyleResolver.ResolveStyles(doc);
            var p = FindFirstElementById(doc.Root, "p");
            var s = FindFirstElementById(doc.Root, "s");
            var e = FindFirstElementById(doc.Root, "e");
            var b = FindFirstElementById(doc.Root, "b");
            Assert.NotNull(p);
            Assert.NotNull(s);
            Assert.NotNull(e);
            Assert.NotNull(b);
            Assert.Equal("#112233", p!.ComputedStyle["color"]);
            Assert.Equal("12px", p.ComputedStyle["margin-top"]);
            Assert.Equal("#112233", s!.ComputedStyle["color"]);
            Assert.False(e!.ComputedStyle.ContainsKey("margin-top"));
            Assert.False(b!.ComputedStyle.ContainsKey("color"));
        }

        [Fact]
        public void ResolveStyles_CurrentColor_ReplacedInDependentValues()
        {
            var doc = HtmlDocument.Parse("<style>p{color:#223344;border-color:currentColor;text-decoration:underline currentColor;}</style><p id=\"p\">x</p>");
            StyleResolver.ResolveStyles(doc);
            var p = FindFirstElementById(doc.Root, "p");
            Assert.NotNull(p);
            Assert.Equal("#223344", p!.ComputedStyle["border-color"]);
            Assert.Contains("#223344", p.ComputedStyle["text-decoration"]);
        }

        [Fact]
        public void ResolveStyles_MinMaxClamp_Basic()
        {
            var doc = HtmlDocument.Parse("<style>p{width:min(300px, 240px);height:max(120px, 200px);margin-left:clamp(8px, 20px, 16px);}</style><p id=\"p\">x</p>");
            StyleResolver.ResolveStyles(doc);
            var p = FindFirstElementById(doc.Root, "p");
            Assert.NotNull(p);
            Assert.Equal("240px", p!.ComputedStyle["width"]);
            Assert.Equal("200px", p.ComputedStyle["height"]);
            Assert.Equal("16px", p.ComputedStyle["margin-left"]);
        }

        [Fact]
        public void ResolveStyles_MinMaxClamp_WithCalc()
        {
            var doc = HtmlDocument.Parse("<style>p{width:min(calc(200px - 20px), 190px);margin-right:clamp(0px, calc(2px + 3px), 10px);}</style><p id=\"p\">x</p>");
            StyleResolver.ResolveStyles(doc);
            var p = FindFirstElementById(doc.Root, "p");
            Assert.NotNull(p);
            Assert.Equal("180px", p!.ComputedStyle["width"]);
            Assert.Equal("5px", p.ComputedStyle["margin-right"]);
        }

        [Fact]
        public void ResolveStyles_MinMaxClamp_MixedUnitsRemainOriginal()
        {
            var doc = HtmlDocument.Parse("<style>p{width:min(100%, 200px);}</style><p id=\"p\">x</p>");
            StyleResolver.ResolveStyles(doc);
            var p = FindFirstElementById(doc.Root, "p");
            Assert.NotNull(p);
            Assert.Equal("min(100%, 200px)", p!.ComputedStyle["width"]);
        }

        [Fact]
        public void MatchesSelector_NthChildWithOfSelector()
        {
            var parent = MakeNode("ul");
            var a = MakeNode("li", @class: "hit");
            var b = MakeNode("li");
            var c = MakeNode("li", @class: "hit");
            a.Parent = parent;
            b.Parent = parent;
            c.Parent = parent;
            parent.Children.Add(a);
            parent.Children.Add(b);
            parent.Children.Add(c);

            Assert.True(InvokeMatches(c, "li:nth-child(2 of .hit)"));
            Assert.False(InvokeMatches(a, "li:nth-child(2 of .hit)"));
            Assert.True(InvokeMatches(a, "li:nth-last-child(2 of .hit)"));
        }

        [Fact]
        public void ResolveStyles_NthChildWithOfSelector()
        {
            var doc = HtmlDocument.Parse("<style>li{color:black;} li:nth-child(2 of .hit){color:red;} li:nth-last-child(1 of .hit){font-weight:bold;}</style><ul><li class=\"hit\" id=\"a\">a</li><li id=\"b\">b</li><li class=\"hit\" id=\"c\">c</li></ul>");
            StyleResolver.ResolveStyles(doc);
            var a = FindFirstElementById(doc.Root, "a");
            var b = FindFirstElementById(doc.Root, "b");
            var c = FindFirstElementById(doc.Root, "c");
            Assert.NotNull(a);
            Assert.NotNull(b);
            Assert.NotNull(c);
            Assert.Equal("black", a!.ComputedStyle["color"]);
            Assert.Equal("black", b!.ComputedStyle["color"]);
            Assert.Equal("red", c!.ComputedStyle["color"]);
            Assert.Equal("bold", c.ComputedStyle["font-weight"]);
        }

        [Fact]
        public void MatchesSelector_NthOfTypeWithOfSelector()
        {
            var parent = MakeNode("div");
            var a = MakeNode("li", @class: "hit");
            var b = MakeNode("li");
            var c = MakeNode("li", @class: "hit");
            var d = MakeNode("span", @class: "hit");
            a.Parent = parent;
            b.Parent = parent;
            c.Parent = parent;
            d.Parent = parent;
            parent.Children.Add(a);
            parent.Children.Add(b);
            parent.Children.Add(c);
            parent.Children.Add(d);

            Assert.True(InvokeMatches(c, "li:nth-of-type(2 of .hit)"));
            Assert.False(InvokeMatches(a, "li:nth-of-type(2 of .hit)"));
            Assert.True(InvokeMatches(a, "li:nth-last-of-type(2 of .hit)"));
        }

        [Fact]
        public void ResolveStyles_NthOfTypeWithOfSelector()
        {
            var doc = HtmlDocument.Parse("<style>li{color:black;} li:nth-of-type(2 of .hit){color:red;} li:nth-last-of-type(1 of .hit){font-weight:bold;}</style><div><li class=\"hit\" id=\"a\">a</li><li id=\"b\">b</li><li class=\"hit\" id=\"c\">c</li><span class=\"hit\" id=\"s\">s</span></div>");
            StyleResolver.ResolveStyles(doc);
            var a = FindFirstElementById(doc.Root, "a");
            var b = FindFirstElementById(doc.Root, "b");
            var c = FindFirstElementById(doc.Root, "c");
            var s = FindFirstElementById(doc.Root, "s");
            Assert.NotNull(a);
            Assert.NotNull(b);
            Assert.NotNull(c);
            Assert.NotNull(s);
            Assert.Equal("black", a!.ComputedStyle["color"]);
            Assert.Equal("black", b!.ComputedStyle["color"]);
            Assert.Equal("red", c!.ComputedStyle["color"]);
            Assert.Equal("bold", c.ComputedStyle["font-weight"]);
            Assert.False(s!.ComputedStyle.ContainsKey("font-weight"));
        }

        [Fact]
        public void ResolveStyles_LogicalProperties_MapToPhysical_Ltr()
        {
            var doc = HtmlDocument.Parse("<style>div{margin-inline-start:10px;margin-inline-end:20px;padding-block-start:3px;padding-block-end:4px;}</style><div id=\"d\">x</div>");
            StyleResolver.ResolveStyles(doc);
            var d = FindFirstElementById(doc.Root, "d");
            Assert.NotNull(d);
            Assert.Equal("10px", d!.ComputedStyle["margin-left"]);
            Assert.Equal("20px", d.ComputedStyle["margin-right"]);
            Assert.Equal("3px", d.ComputedStyle["padding-top"]);
            Assert.Equal("4px", d.ComputedStyle["padding-bottom"]);
        }

        [Fact]
        public void ResolveStyles_LogicalProperties_MapToPhysical_Rtl()
        {
            var doc = HtmlDocument.Parse("<style>div{margin-inline:1px 2px;padding-inline-start:5px;padding-inline-end:6px;}</style><div dir=\"rtl\" id=\"d\">x</div>");
            StyleResolver.ResolveStyles(doc);
            var d = FindFirstElementById(doc.Root, "d");
            Assert.NotNull(d);
            Assert.Equal("1px", d!.ComputedStyle["margin-right"]);
            Assert.Equal("2px", d.ComputedStyle["margin-left"]);
            Assert.Equal("5px", d.ComputedStyle["padding-right"]);
            Assert.Equal("6px", d.ComputedStyle["padding-left"]);
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

        private static HtmlNode? FindFirstElementById(HtmlNode node, string id)
        {
            if (node.GetAttribute("id")?.Equals(id, StringComparison.OrdinalIgnoreCase) == true)
                return node;

            foreach (var child in node.Children)
            {
                var found = FindFirstElementById(child, id);
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
