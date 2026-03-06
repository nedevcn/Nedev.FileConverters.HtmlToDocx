using System;
using System.Collections.Generic;
using System.Linq;

namespace Nedev.FileConverters.HtmlToDocx.Core.Css;

using Nedev.FileConverters.HtmlToDocx.Core.Html;

public sealed class StyleResolver
{
    private static readonly HashSet<string> InheritableProperties = new(StringComparer.OrdinalIgnoreCase)
    {
        "font-family", "font-size", "font-weight", "font-style", "color", "text-align", "line-height"
    };

    public static void ResolveStyles(HtmlDocument document)
    {
        ResolveNodeStyle(document.Root, document.Stylesheet, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));
    }

    private static void ResolveNodeStyle(HtmlNode node, List<CssRule> stylesheet, Dictionary<string, string> parentStyle)
    {
        var computed = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        // 1. Inheritance
        foreach (var kvp in parentStyle)
        {
            if (InheritableProperties.Contains(kvp.Key))
            {
                computed[kvp.Key] = kvp.Value;
            }
        }

        // 2. Stylesheet Rules (Simplified matching)
        foreach (var rule in stylesheet)
        {
            if (MatchesSelector(node, rule.Selector))
            {
                foreach (var decl in rule.Declarations)
                {
                    computed[decl.Key] = decl.Value;
                }
            }
        }

        // 3. Inline Styles (Highest priority)
        var inlineStyleAttr = node.GetAttribute("style");
        if (!string.IsNullOrEmpty(inlineStyleAttr))
        {
            var inlineDecls = CssParser.ParseInlineStyle(inlineStyleAttr);
            foreach (var decl in inlineDecls)
            {
                computed[decl.Key] = decl.Value;
            }
        }

        node.ComputedStyle = computed;

        foreach (var child in node.Children)
        {
            ResolveNodeStyle(child, stylesheet, computed);
        }
    }

    private static bool MatchesSelector(HtmlNode node, string selector)
    {
        if (string.IsNullOrWhiteSpace(selector)) return false;

        // split into tokens (includes '>' when used)
        var tokens = selector.Split(' ', StringSplitOptions.RemoveEmptyEntries);
#if DEBUG
        Console.WriteLine($"[DEBUG] MatchesSelector('{selector}') -> tokens: {string.Join("|", tokens)}");
#endif
        var result = MatchSelectorParts(node, tokens, tokens.Length - 1);
#if DEBUG
        Console.WriteLine($"[DEBUG] MatchesSelector result = {result}");
#endif
        return result;
    }

    // recursive matcher that understands descendant and child combinators
    private static bool MatchSelectorParts(HtmlNode? node, string[] parts, int index)
    {
        if (index < 0) return true;
        if (node == null) return false;

        var part = parts[index];
        if (part == ">")
        {
            // '>' should be handled as part of previous recursion step
            return false;
        }

        if (!MatchesSimple(node, part))
            return false;

        if (index == 0) // no more selectors to match
            return true;

        // check for explicit child combinator before the next token
        if (parts[index - 1] == ">")
        {
            // direct parent must match next selector
            return MatchSelectorParts(node.Parent, parts, index - 2);
        }

        // descendant combinator: search ancestors for a match
        var ancestor = node.Parent;
        while (ancestor != null)
        {
            if (MatchSelectorParts(ancestor, parts, index - 1))
                return true;
            ancestor = ancestor.Parent;
        }

        return false;
    }

    // simple selector matching (tag, #id, .class combinations)
    private static bool MatchesSimple(HtmlNode node, string sel)
    {
        if (string.IsNullOrEmpty(sel)) return false;

        // parse tag, id and classes from selector string
        string tag = sel;
        string? id = null;
        var classes = new List<string>();

        int pos = 0;
        while (pos < tag.Length)
        {
            if (tag[pos] == '#')
            {
                id = tag.Substring(pos + 1);
                tag = tag.Substring(0, pos);
                break;
            }
            if (tag[pos] == '.')
            {
                // gather all class segments
                var parts = tag.Substring(pos + 1).Split('.', StringSplitOptions.RemoveEmptyEntries);
                classes.AddRange(parts);
                tag = tag.Substring(0, pos);
                break;
            }
            pos++;
        }

        if (!string.IsNullOrEmpty(tag) && !node.TagName.Equals(tag, StringComparison.OrdinalIgnoreCase))
            return false;
        if (id != null && !id.Equals(node.Id, StringComparison.OrdinalIgnoreCase))
            return false;
        if (classes.Count > 0)
        {
            var nodeClasses = node.GetAttribute("class")?.Split(' ', StringSplitOptions.RemoveEmptyEntries) ?? new string[0];
            foreach (var cls in classes)
                if (!nodeClasses.Any(c => c.Equals(cls, StringComparison.OrdinalIgnoreCase)))
                    return false;
        }

        return true;
    }
}
