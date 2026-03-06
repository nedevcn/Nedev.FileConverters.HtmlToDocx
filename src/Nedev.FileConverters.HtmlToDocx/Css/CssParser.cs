using System;
using System.Collections.Generic;

namespace Nedev.FileConverters.HtmlToDocx.Core.Css;

public sealed class CssParser
{
    public static Dictionary<string, string> ParseInlineStyle(string? style)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(style)) return result;

        var span = style.AsSpan();
        int start = 0;
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

            if (!string.IsNullOrEmpty(name))
            {
                result[name] = value;
            }
        }

        return result;
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
                rules.Add(new CssRule(sel.Trim(), ParseInlineStyle(body)));
            }

            pos += braceOpen + braceClose + 1;
        }

        return rules;
    }
}

public sealed class CssRule
{
    public string Selector { get; }
    public Dictionary<string, string> Declarations { get; }

    public CssRule(string selector, Dictionary<string, string> declarations)
    {
        Selector = selector;
        Declarations = declarations;
    }
}
