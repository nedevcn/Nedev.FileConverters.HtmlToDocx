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

        // Basic selectors: tag, .class, #id
        if (selector.StartsWith("#"))
        {
            return node.Id?.Equals(selector.Substring(1), StringComparison.OrdinalIgnoreCase) ?? false;
        }
        
        if (selector.StartsWith("."))
        {
            var className = node.GetAttribute("class");
            if (string.IsNullOrEmpty(className)) return false;
            var targetClass = selector.Substring(1);
            return className.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                            .Any(c => c.Equals(targetClass, StringComparison.OrdinalIgnoreCase));
        }

        return node.TagName.Equals(selector, StringComparison.OrdinalIgnoreCase);
    }
}
