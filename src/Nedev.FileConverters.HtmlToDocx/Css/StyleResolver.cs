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

    private static readonly HashSet<string> Combinators = new(StringComparer.Ordinal)
    {
        " ", ">", "+", "~"
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

        var tokens = TokenizeSelector(selector);
        if (tokens.Count == 0) return false;

        // last token must be a simple selector
        if (Combinators.Contains(tokens[^1])) return false;

        return MatchTokensFromRight(node, tokens, tokens.Count - 1);
    }

    // Kept for tests (reflection) and compatibility. This previously expected an array from selector.Split(' ').
    // We now normalize it into a full token stream (inserting descendant combinators between adjacent simples).
    private static bool MatchSelectorParts(HtmlNode? node, string[] parts, int index)
    {
        if (node == null) return false;
        if (parts == null || parts.Length == 0) return false;
        if (index < 0) return true;

        // Build tokens from the slice [0..index]
        var tokens = new List<string>(capacity: index + 1);
        for (int i = 0; i <= index && i < parts.Length; i++)
        {
            var p = parts[i];
            if (string.IsNullOrWhiteSpace(p)) continue;

            var isComb = Combinators.Contains(p);
            if (!isComb)
            {
                // insert descendant combinator if previous token is also a simple selector
                if (tokens.Count > 0 && !Combinators.Contains(tokens[^1]))
                    tokens.Add(" ");
                tokens.Add(p);
            }
            else
            {
                // ignore invalid leading combinators; otherwise replace any previous combinator
                if (tokens.Count == 0) continue;
                if (Combinators.Contains(tokens[^1])) tokens[^1] = p;
                else tokens.Add(p);
            }
        }

        if (tokens.Count == 0) return false;
        if (Combinators.Contains(tokens[^1])) return false;

        return MatchTokensFromRight(node, tokens, tokens.Count - 1);
    }

    private static bool MatchTokensFromRight(HtmlNode? node, IReadOnlyList<string> tokens, int index)
    {
        if (index < 0) return true;
        if (node == null) return false;

        var simple = tokens[index];
        if (Combinators.Contains(simple)) return false;
        if (!MatchesSimple(node, simple)) return false;

        if (index == 0) return true;

        // tokens are "... <simple> <combinator> <simple>" so combinator is index-1, left simple is index-2
        var combinator = tokens[index - 1];
        if (!Combinators.Contains(combinator)) combinator = " ";

        if (index - 2 < 0) return true;

        return combinator switch
        {
            ">" => MatchTokensFromRight(node.Parent, tokens, index - 2),
            " " => MatchDescendant(node, tokens, index - 2),
            "+" => MatchAdjacentSibling(node, tokens, index - 2),
            "~" => MatchGeneralSibling(node, tokens, index - 2),
            _ => MatchDescendant(node, tokens, index - 2)
        };
    }

    private static bool MatchDescendant(HtmlNode node, IReadOnlyList<string> tokens, int leftIndex)
    {
        var ancestor = node.Parent;
        while (ancestor != null)
        {
            if (MatchTokensFromRight(ancestor, tokens, leftIndex)) return true;
            ancestor = ancestor.Parent;
        }
        return false;
    }

    private static bool MatchAdjacentSibling(HtmlNode node, IReadOnlyList<string> tokens, int leftIndex)
    {
        var prev = GetPreviousElementSibling(node);
        return prev != null && MatchTokensFromRight(prev, tokens, leftIndex);
    }

    private static bool MatchGeneralSibling(HtmlNode node, IReadOnlyList<string> tokens, int leftIndex)
    {
        var parent = node.Parent;
        if (parent == null) return false;

        var siblings = parent.Children;
        var pos = siblings.IndexOf(node);
        if (pos <= 0) return false;

        for (int i = pos - 1; i >= 0; i--)
        {
            var sib = siblings[i];
            if (sib.TagName == "#text") continue;
            if (MatchTokensFromRight(sib, tokens, leftIndex)) return true;
        }

        return false;
    }

    private static HtmlNode? GetPreviousElementSibling(HtmlNode node)
    {
        var parent = node.Parent;
        if (parent == null) return null;

        var siblings = parent.Children;
        var pos = siblings.IndexOf(node);
        if (pos <= 0) return null;

        for (int i = pos - 1; i >= 0; i--)
        {
            var sib = siblings[i];
            if (sib.TagName == "#text") continue;
            return sib;
        }

        return null;
    }

    private static List<string> TokenizeSelector(string selector)
    {
        var tokens = new List<string>();
        int i = 0;

        while (i < selector.Length)
        {
            // consume whitespace as descendant combinator, but only between two simple selectors
            if (char.IsWhiteSpace(selector[i]))
            {
                while (i < selector.Length && char.IsWhiteSpace(selector[i])) i++;
                if (tokens.Count > 0 && !Combinators.Contains(tokens[^1]))
                    tokens.Add(" ");
                continue;
            }

            char c = selector[i];
            if (c == '>' || c == '+' || c == '~')
            {
                if (tokens.Count > 0 && Combinators.Contains(tokens[^1])) tokens[^1] = c.ToString();
                else tokens.Add(c.ToString());
                i++;
                continue;
            }

            int start = i;
            int bracketDepth = 0;
            char quote = '\0';
            while (i < selector.Length)
            {
                c = selector[i];
                if (quote != '\0')
                {
                    if (c == quote) quote = '\0';
                    i++;
                    continue;
                }

                if (c == '"' || c == '\'')
                {
                    quote = c;
                    i++;
                    continue;
                }

                if (c == '[') { bracketDepth++; i++; continue; }
                if (c == ']') { if (bracketDepth > 0) bracketDepth--; i++; continue; }

                if (bracketDepth == 0 && (char.IsWhiteSpace(c) || c == '>' || c == '+' || c == '~'))
                    break;

                i++;
            }

            var simple = selector.Substring(start, i - start).Trim();
            if (!string.IsNullOrEmpty(simple))
            {
                if (tokens.Count > 0 && !Combinators.Contains(tokens[^1]))
                    tokens.Add(" ");
                tokens.Add(simple);
            }
        }

        // remove trailing combinators
        while (tokens.Count > 0 && Combinators.Contains(tokens[^1])) tokens.RemoveAt(tokens.Count - 1);

        return tokens;
    }

    // simple selector matching (tag, #id, .class combinations)
    private static bool MatchesSimple(HtmlNode node, string sel)
    {
        if (string.IsNullOrEmpty(sel)) return false;

        string tag = string.Empty;
        string? id = null;
        var classes = new List<string>();
        var attributes = new List<(string name, string? value)>();

        int i = 0;
        while (i < sel.Length)
        {
            char c = sel[i];
            if (c == '#')
            {
                i++;
                int start = i;
                while (i < sel.Length && sel[i] != '.' && sel[i] != '#' && sel[i] != '[') i++;
                id = sel[start..i];
            }
            else if (c == '.')
            {
                i++;
                int start = i;
                while (i < sel.Length && sel[i] != '.' && sel[i] != '#' && sel[i] != '[') i++;
                classes.Add(sel[start..i]);
            }
            else if (c == '[')
            {
                i++;
                int start = i;
                while (i < sel.Length && sel[i] != ']') i++;
                var inside = sel[start..i];
                i++; // skip ']'
                string name = inside;
                string? val = null;
                var eq = inside.IndexOf('=');
                if (eq >= 0)
                {
                    name = inside.Substring(0, eq);
                    val = inside.Substring(eq + 1).Trim('"', '\'');
                }
                attributes.Add((name, val));
            }
            else
            {
                // tag name
                int start = i;
                while (i < sel.Length && sel[i] != '.' && sel[i] != '#' && sel[i] != '[') i++;
                tag = sel[start..i];
            }
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
        if (attributes.Count > 0)
        {
            foreach (var (name, val) in attributes)
            {
                var attr = node.GetAttribute(name);
                if (attr == null) return false;
                if (val != null && !attr.Equals(val, StringComparison.OrdinalIgnoreCase))
                    return false;
            }
        }

        return true;
    }
}
