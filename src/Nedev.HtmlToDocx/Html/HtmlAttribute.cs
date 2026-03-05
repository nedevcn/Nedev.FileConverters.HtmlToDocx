namespace Nedev.HtmlToDocx.Core.Html;

public readonly struct HtmlAttribute
{
    public string Name { get; }
    public string Value { get; }

    public HtmlAttribute(string name, string? value)
    {
        Name = name ?? string.Empty;
        Value = value ?? string.Empty;
    }
}
