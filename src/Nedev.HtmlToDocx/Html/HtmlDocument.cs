namespace Nedev.HtmlToDocx.Core.Html;

public class HtmlNode
{
    public string TagName { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
    public List<HtmlNode> Children { get; set; } = new();
    public HtmlNode? Parent { get; set; }
    public Dictionary<string, string> Attributes { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public string? GetAttribute(string name) => Attributes.TryGetValue(name, out var value) ? value : null;
    public string? Id => GetAttribute("id");
}

public class HtmlDocument
{
    public HtmlNode Root { get; } = new() { TagName = "#document" };

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
                        var textNode = new HtmlNode { TagName = "#text", Text = token.Text, Parent = stack.Peek() };
                        stack.Peek().Children.Add(textNode);
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
