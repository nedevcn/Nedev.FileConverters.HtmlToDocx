using System;
using System.Collections.Generic;

namespace Nedev.FileConverters.HtmlToDocx.Core.Css;

public sealed class CssParser
{
    public static Dictionary<string, string> ParseInlineStyle(string? style)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(style)) return result;

        var declarations = ParseDeclarations(style);
        foreach (var declaration in declarations)
            result[declaration.Property] = declaration.Value;

        return result;
    }

    public static List<CssDeclaration> ParseDeclarations(string? style)
    {
        var declarations = new List<CssDeclaration>();
        if (string.IsNullOrWhiteSpace(style)) return declarations;

        var span = style.AsSpan();
        int start = 0;
        int sourceOrder = 0;

        while (start < span.Length)
        {
            int colonIndex = span.Slice(start).IndexOf(':');
            if (colonIndex == -1) break;

            var name = span.Slice(start, colonIndex).Trim().ToString().ToLowerInvariant();
            int valueStart = start + colonIndex + 1;
            int semicolonIndex = span.Slice(valueStart).IndexOf(';');

            string value;
            if (semicolonIndex == -1)
            {
                value = span.Slice(valueStart).Trim().ToString();
                start = span.Length;
            }
            else
            {
                value = span.Slice(valueStart, semicolonIndex).Trim().ToString();
                start = valueStart + semicolonIndex + 1;
            }

            if (string.IsNullOrEmpty(name)) continue;

            var important = TryStripImportant(ref value);
            declarations.Add(new CssDeclaration(name, value, important, sourceOrder++));
        }

        return declarations;
    }

    public static List<CssRule> ParseStylesheet(string css)
    {
        var rules = new List<CssRule>();
        if (string.IsNullOrWhiteSpace(css)) return rules;

        var span = css.AsSpan();
        int pos = 0;
        while (pos < span.Length)
        {
            // Skip whitespace and comments
            if (char.IsWhiteSpace(span[pos])) { pos++; continue; }
            if (pos + 1 < span.Length && span[pos] == '/' && span[pos+1] == '*')
            {
                int endComment = span.Slice(pos).IndexOf("*/", StringComparison.Ordinal);
                if (endComment == -1) break;
                pos += endComment + 2;
                continue;
            }

            int braceOpen = span.Slice(pos).IndexOf('{');
            if (braceOpen == -1) break;

            // selector list may be comma-separated
            var rawSelector = span.Slice(pos, braceOpen).Trim().ToString();
            int braceClose = span.Slice(pos + braceOpen).IndexOf('}');
            if (braceClose == -1) break;

            var body = span.Slice(pos + braceOpen + 1, braceClose - 1).ToString();

            // break selectors by comma and create a rule per selector
            var selectors = rawSelector.Split(',', StringSplitOptions.RemoveEmptyEntries);
            foreach (var sel in selectors)
            {
                rules.Add(new CssRule(sel.Trim(), ParseDeclarations(body)));
            }

            pos += braceOpen + braceClose + 1;
        }

        return rules;
    }

    private static bool TryStripImportant(ref string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;

        int marker = value.LastIndexOf('!');
        if (marker < 0) return false;

        var suffix = value.Substring(marker).Trim();
        if (!suffix.Equals("!important", StringComparison.OrdinalIgnoreCase)) return false;

        value = value.Substring(0, marker).TrimEnd();
        return true;
    }
}

public sealed class CssRule
{
    public string Selector { get; }
    public List<CssDeclaration> Declarations { get; }

    public CssRule(string selector, List<CssDeclaration> declarations)
    {
        Selector = selector;
        Declarations = declarations;
    }
}

public sealed class CssDeclaration
{
    public string Property { get; }
    public string Value { get; }
    public bool Important { get; }
    public int SourceOrder { get; }

    public CssDeclaration(string property, string value, bool important, int sourceOrder)
    {
        Property = property;
        Value = value;
        Important = important;
        SourceOrder = sourceOrder;
    }
}
