using System.Collections.Generic;
using Nedev.FileConverters.HtmlToDocx.Core.Css;

namespace Nedev.FileConverters.HtmlToDocx.Core.Html;

public class HtmlNode
{
    public string TagName { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public List<HtmlNode> Children { get; set; } = new();
    public HtmlNode? Parent { get; set; }
    public Dictionary<string, string> Attributes { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, string> ComputedStyle { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public string? GetAttribute(string name) => Attributes.TryGetValue(name, out var value) ? value : null;
    public string? Id => GetAttribute("id");
}

public class HtmlDocument
{
    public HtmlNode Root { get; } = new() { TagName = "#document" };
    public List<CssRule> Stylesheet { get; } = new();

    private static readonly HashSet<string> VoidTags = new(StringComparer.OrdinalIgnoreCase)
    {
        "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr"
    };

    private static readonly HashSet<string> PreserveWhitespaceParents = new(StringComparer.OrdinalIgnoreCase)
    {
        // Inline-ish / text container elements where whitespace-only nodes can matter
        "p", "span", "a", "b", "strong", "i", "em", "u", "font", "li",
        "td", "th",
        "h1", "h2", "h3", "h4", "h5", "h6"
    };

    private static readonly HashSet<string> PreformattedTags = new(StringComparer.OrdinalIgnoreCase)
    {
        "pre", "code", "textarea"
    };

    private static readonly HashSet<string> BlockTags = new(StringComparer.OrdinalIgnoreCase)
    {
        "address", "article", "aside", "blockquote", "div", "dl", "dt", "dd", "fieldset", "figcaption", "figure", "footer",
        "form", "h1", "h2", "h3", "h4", "h5", "h6", "header", "hr", "li", "main", "nav", "ol", "p", "pre", "section",
        "table", "thead", "tbody", "tfoot", "tr", "td", "th", "ul"
    };

    public static HtmlDocument BuildTree(HtmlToken[] tokens)
    {
        var document = new HtmlDocument();
        var stack = new Stack<HtmlNode>();
        stack.Push(document.Root);

        foreach (var token in tokens)
        {
            switch (token.Type)
            {
                case HtmlTokenType.Text:
                    if (!string.IsNullOrEmpty(token.Text))
                    {
                        var parent = stack.Peek();
                        // Special handling for style tag content
                        if (parent.TagName.Equals("style", StringComparison.OrdinalIgnoreCase))
                        {
                            document.Stylesheet.AddRange(Css.CssParser.ParseStylesheet(token.Text));
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(token.Text))
                            {
                                // Keep a single space only when the parent is a text-bearing element.
                                if (PreserveWhitespaceParents.Contains(parent.TagName))
                                {
                                    AppendTextNode(parent, " ");
                                }
                            }
                            else
                            {
                                var text = InPreformattedContext(parent) ? token.Text : NormalizeWhitespace(token.Text);
                                if (!string.IsNullOrEmpty(text))
                                    AppendTextNode(parent, text);
                            }
                        }
                    }
                    break;
                case HtmlTokenType.StartTag:
                    // Basic error-tolerance similar to HTML parser behavior:
                    // - implicit closes for some elements (e.g. <li><li>, <p><div>)
                    ApplyImplicitCloseRules(stack, token.TagName);

                    var element = new HtmlNode { TagName = token.TagName, Parent = stack.Peek() };
                    foreach (var attr in token.Attributes)
                        element.Attributes[attr.Name] = attr.Value;
                    stack.Peek().Children.Add(element);
                    // Don't push void elements even if tokenizer mis-classified them as StartTag.
                    if (!VoidTags.Contains(token.TagName))
                        stack.Push(element);
                    break;
                case HtmlTokenType.EndTag:
                    if (VoidTags.Contains(token.TagName))
                        break;

                    // Pop until matching tag is found (tolerates mismatched nesting).
                    while (stack.Count > 1)
                    {
                        if (stack.Peek().TagName.Equals(token.TagName, StringComparison.OrdinalIgnoreCase))
                        {
                            stack.Pop();
                            break;
                        }
                        stack.Pop();
                    }
                    break;
                case HtmlTokenType.SelfClosingTag:
                    var selfClosing = new HtmlNode { TagName = token.TagName, Parent = stack.Peek() };
                    foreach (var attr in token.Attributes)
                        selfClosing.Attributes[attr.Name] = attr.Value;
                    stack.Peek().Children.Add(selfClosing);
                    break;
                case HtmlTokenType.Comment:
                case HtmlTokenType.Doctype:
                    // For now, we ignore these in the tree
                    break;
            }
        }
        return document;
    }

    public static HtmlDocument Parse(string html)
    {
        var parser = new HtmlParser(html.AsSpan());
        var tokens = parser.ParseAll();
        return BuildTree(tokens);
    }

    private static void AppendTextNode(HtmlNode parent, string text)
    {
        if (string.IsNullOrEmpty(text)) return;

        if (parent.Children.Count > 0 && parent.Children[^1].TagName == "#text")
        {
            // Merge adjacent text nodes to preserve spaces across tag boundaries.
            var last = parent.Children[^1];
            var combined = last.Text + text;
            last.Text = InPreformattedContext(parent) ? combined : NormalizeWhitespace(combined);
            return;
        }

        parent.Children.Add(new HtmlNode { TagName = "#text", Text = text, Parent = parent });
    }

    private static bool InPreformattedContext(HtmlNode node)
    {
        HtmlNode? current = node;
        while (current != null)
        {
            if (PreformattedTags.Contains(current.TagName))
                return true;
            current = current.Parent;
        }
        return false;
    }

    private static string NormalizeWhitespace(string text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;

        // Collapse any whitespace run into a single space, preserving leading/trailing as at most one space.
        var sb = new System.Text.StringBuilder(text.Length);
        bool lastWasSpace = false;

        foreach (var ch in text)
        {
            if (char.IsWhiteSpace(ch))
            {
                if (!lastWasSpace)
                {
                    sb.Append(' ');
                    lastWasSpace = true;
                }
            }
            else
            {
                sb.Append(ch);
                lastWasSpace = false;
            }
        }

        return sb.ToString();
    }

    private static void ApplyImplicitCloseRules(Stack<HtmlNode> stack, string newTagName)
    {
        if (stack.Count <= 1) return;

        var top = stack.Peek();
        if (top.TagName == "#document") return;

        // Table fixups for common omitted tags:
        // - <td>/<th> implicitly close on next cell or next row
        // - <tr> implicitly closes on next row or when a section starts
        // - <thead>/<tbody>/<tfoot> implicitly close on the next section
        if (newTagName.Equals("tr", StringComparison.OrdinalIgnoreCase))
        {
            // If a cell is still open, close it before starting a new row.
            while (stack.Count > 1)
            {
                top = stack.Peek();
                if (top.TagName.Equals("td", StringComparison.OrdinalIgnoreCase) ||
                    top.TagName.Equals("th", StringComparison.OrdinalIgnoreCase))
                {
                    stack.Pop();
                    continue;
                }
                break;
            }

            top = stack.Peek();
            if (top.TagName.Equals("tr", StringComparison.OrdinalIgnoreCase))
            {
                stack.Pop();
                return;
            }
        }

        if (newTagName.Equals("td", StringComparison.OrdinalIgnoreCase) ||
            newTagName.Equals("th", StringComparison.OrdinalIgnoreCase))
        {
            // Close previous cell if still open.
            if (top.TagName.Equals("td", StringComparison.OrdinalIgnoreCase) ||
                top.TagName.Equals("th", StringComparison.OrdinalIgnoreCase))
            {
                stack.Pop();
                return;
            }

            // If we're not in a row, but inside table-ish nodes, pop up to row.
            if (!top.TagName.Equals("tr", StringComparison.OrdinalIgnoreCase))
            {
                while (stack.Count > 1)
                {
                    top = stack.Peek();
                    if (top.TagName.Equals("tr", StringComparison.OrdinalIgnoreCase)) break;
                    if (top.TagName.Equals("table", StringComparison.OrdinalIgnoreCase) ||
                        top.TagName.Equals("thead", StringComparison.OrdinalIgnoreCase) ||
                        top.TagName.Equals("tbody", StringComparison.OrdinalIgnoreCase) ||
                        top.TagName.Equals("tfoot", StringComparison.OrdinalIgnoreCase))
                    {
                        break;
                    }
                    stack.Pop();
                }
            }
        }

        if (newTagName.Equals("thead", StringComparison.OrdinalIgnoreCase) ||
            newTagName.Equals("tbody", StringComparison.OrdinalIgnoreCase) ||
            newTagName.Equals("tfoot", StringComparison.OrdinalIgnoreCase))
        {
            // Close any open row/cell before starting a new section.
            while (stack.Count > 1)
            {
                top = stack.Peek();
                if (top.TagName.Equals("td", StringComparison.OrdinalIgnoreCase) ||
                    top.TagName.Equals("th", StringComparison.OrdinalIgnoreCase) ||
                    top.TagName.Equals("tr", StringComparison.OrdinalIgnoreCase))
                {
                    stack.Pop();
                    continue;
                }
                break;
            }

            top = stack.Peek();
            if (top.TagName.Equals("thead", StringComparison.OrdinalIgnoreCase) ||
                top.TagName.Equals("tbody", StringComparison.OrdinalIgnoreCase) ||
                top.TagName.Equals("tfoot", StringComparison.OrdinalIgnoreCase))
            {
                stack.Pop();
                return;
            }
        }

        // <p> is implicitly closed by many block elements.
        if (top.TagName.Equals("p", StringComparison.OrdinalIgnoreCase) && BlockTags.Contains(newTagName) && !newTagName.Equals("span", StringComparison.OrdinalIgnoreCase))
        {
            stack.Pop();
            return;
        }

        // These elements auto-close when the same element starts again.
        if (top.TagName.Equals(newTagName, StringComparison.OrdinalIgnoreCase))
        {
            if (newTagName.Equals("li", StringComparison.OrdinalIgnoreCase) ||
                newTagName.Equals("tr", StringComparison.OrdinalIgnoreCase) ||
                newTagName.Equals("td", StringComparison.OrdinalIgnoreCase) ||
                newTagName.Equals("th", StringComparison.OrdinalIgnoreCase) ||
                newTagName.Equals("p", StringComparison.OrdinalIgnoreCase))
            {
                stack.Pop();
            }
        }

        // <td> and <th> close each other on new cell start
        if ((newTagName.Equals("td", StringComparison.OrdinalIgnoreCase) || newTagName.Equals("th", StringComparison.OrdinalIgnoreCase)) &&
            (top.TagName.Equals("td", StringComparison.OrdinalIgnoreCase) || top.TagName.Equals("th", StringComparison.OrdinalIgnoreCase)))
        {
            stack.Pop();
        }
    }
}
