namespace Nedev.HtmlToDocx.Core.Html;

public enum HtmlTokenType
{
    Text,
    StartTag,
    EndTag,
    SelfClosingTag,
    Comment,
    Doctype,
    EOF
}

public readonly struct HtmlToken
{
    public HtmlTokenType Type { get; }
    public string TagName { get; }
    public string Text { get; }
    public HtmlAttribute[] Attributes { get; }
    public bool IsSelfClosing { get; }

    public HtmlToken(HtmlTokenType type, string tagName = "", string text = "", 
        HtmlAttribute[]? attributes = null, bool isSelfClosing = false)
    {
        Type = type;
        TagName = tagName;
        Text = text ?? string.Empty;
        Attributes = attributes ?? Array.Empty<HtmlAttribute>();
        IsSelfClosing = isSelfClosing;
    }

    public static HtmlToken TextToken(string text) => new(HtmlTokenType.Text, text: text);
    public static HtmlToken StartTagToken(string tagName, HtmlAttribute[]? attributes = null) => 
        new(HtmlTokenType.StartTag, tagName: tagName, attributes: attributes);
    public static HtmlToken EndTagToken(string tagName) => new(HtmlTokenType.EndTag, tagName: tagName);
    public static HtmlToken SelfClosingTagToken(string tagName, HtmlAttribute[]? attributes = null) => 
        new(HtmlTokenType.SelfClosingTag, tagName: tagName, attributes: attributes, isSelfClosing: true);
    public static HtmlToken EofToken() => new(HtmlTokenType.EOF);
}
