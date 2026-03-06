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
                    if (!string.IsNullOrWhiteSpace(token.Text))
                    {
                        var parent = stack.Peek();
                        // Special handling for style tag content
                        if (parent.TagName.Equals("style", StringComparison.OrdinalIgnoreCase))
                        {
                            document.Stylesheet.AddRange(Css.CssParser.ParseStylesheet(token.Text));
                        }
                        else
                        {
                            var textNode = new HtmlNode { TagName = "#text", Text = token.Text, Parent = parent };
                            parent.Children.Add(textNode);
                        }
                    }
                    break;
                case HtmlTokenType.StartTag:
                    var element = new HtmlNode { TagName = token.TagName, Parent = stack.Peek() };
                    foreach (var attr in token.Attributes)
                        element.Attributes[attr.Name] = attr.Value;
                    stack.Peek().Children.Add(element);
                    stack.Push(element);
                    break;
                case HtmlTokenType.EndTag:
                    if (stack.Count > 1 && stack.Peek().TagName.Equals(token.TagName, StringComparison.OrdinalIgnoreCase))
                        stack.Pop();
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
}
