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
        var priority = new Dictionary<string, StylePriority>(StringComparer.OrdinalIgnoreCase);

        foreach (var kvp in parentStyle)
        {
            if (InheritableProperties.Contains(kvp.Key))
            {
                computed[kvp.Key] = kvp.Value;
                priority[kvp.Key] = new StylePriority(important: false, specificity: 0, sourceOrder: -1);
            }
        }

        int sourceOrder = 0;
        foreach (var rule in stylesheet)
        {
            if (MatchesSelector(node, rule.Selector))
            {
                int specificity = CalculateSelectorSpecificity(rule.Selector);
                foreach (var decl in rule.Declarations)
                {
                    int declarationOrder = unchecked(sourceOrder * 1000 + decl.SourceOrder);
                    ApplyDeclaration(computed, priority, decl.Property, decl.Value, decl.Important, specificity, declarationOrder);
                }
            }
            sourceOrder++;
        }

        var inlineStyleAttr = node.GetAttribute("style");
        if (!string.IsNullOrEmpty(inlineStyleAttr))
        {
            var inlineDecls = CssParser.ParseDeclarations(inlineStyleAttr);
            foreach (var decl in inlineDecls)
            {
                ApplyDeclaration(
                    computed,
                    priority,
                    decl.Property,
                    decl.Value,
                    decl.Important,
                    specificity: 1000,
                    sourceOrder: int.MaxValue - inlineDecls.Count + decl.SourceOrder
                );
            }
        }

        node.ComputedStyle = computed;

        foreach (var child in node.Children)
        {
            ResolveNodeStyle(child, stylesheet, computed);
        }
    }

    private static void ApplyDeclaration(
        Dictionary<string, string> computed,
        Dictionary<string, StylePriority> priority,
        string property,
        string value,
        bool important,
        int specificity,
        int sourceOrder)
    {
        var incoming = new StylePriority(important, specificity, sourceOrder);
        if (!priority.TryGetValue(property, out var existing) || incoming.CompareTo(existing) >= 0)
        {
            computed[property] = value;
            priority[property] = incoming;
        }
    }

    private static int CalculateSelectorSpecificity(string selector)
    {
        int idCount = 0;
        int classLikeCount = 0;
        int typeCount = 0;
        var tokens = TokenizeSelector(selector);

        foreach (var token in tokens)
        {
            if (Combinators.Contains(token)) continue;
            if (token == "*") continue;
            var (a, b, c) = CalculateSimpleSpecificity(token);
            idCount += a;
            classLikeCount += b;
            typeCount += c;
        }

        return idCount * 100 + classLikeCount * 10 + typeCount;
    }

    private static (int id, int classLike, int type) CalculateSimpleSpecificity(string simple)
    {
        int idCount = 0;
        int classLikeCount = 0;
        int typeCount = 0;
        int i = 0;
        bool hasType = false;

        while (i < simple.Length)
        {
            char c = simple[i];
            if (c == '#')
            {
                idCount++;
                i++;
                while (i < simple.Length && simple[i] != '.' && simple[i] != '#' && simple[i] != '[' && simple[i] != ':') i++;
            }
            else if (c == '.')
            {
                classLikeCount++;
                i++;
                while (i < simple.Length && simple[i] != '.' && simple[i] != '#' && simple[i] != '[' && simple[i] != ':') i++;
            }
            else if (c == '[')
            {
                classLikeCount++;
                i++;
                while (i < simple.Length && simple[i] != ']') i++;
                if (i < simple.Length) i++;
            }
            else if (c == ':')
            {
                i++;
                bool pseudoElement = i < simple.Length && simple[i] == ':';
                if (pseudoElement)
                {
                    typeCount++;
                    i++;
                }

                int nameStart = i;
                while (i < simple.Length && (char.IsLetterOrDigit(simple[i]) || simple[i] == '-' || simple[i] == '_')) i++;
                var pseudoName = simple.Substring(nameStart, i - nameStart);

                if (!pseudoElement && pseudoName.Equals("not", StringComparison.OrdinalIgnoreCase) && i < simple.Length && simple[i] == '(')
                {
                    if (TryReadFunctionArgument(simple, ref i, out var arg))
                    {
                        int maxA = 0, maxB = 0, maxC = 0;
                        foreach (var part in SplitSelectorList(arg))
                        {
                            var trimmed = part.Trim();
                            if (string.IsNullOrEmpty(trimmed)) continue;
                            var specificity = CalculateSelectorSpecificity(trimmed);
                            int a = specificity / 100;
                            int b = (specificity % 100) / 10;
                            int c2 = specificity % 10;
                            if (a > maxA || (a == maxA && b > maxB) || (a == maxA && b == maxB && c2 > maxC))
                            {
                                maxA = a;
                                maxB = b;
                                maxC = c2;
                            }
                        }
                        idCount += maxA;
                        classLikeCount += maxB;
                        typeCount += maxC;
                    }
                }
                else if (!pseudoElement && pseudoName.Equals("is", StringComparison.OrdinalIgnoreCase) && i < simple.Length && simple[i] == '(')
                {
                    if (TryReadFunctionArgument(simple, ref i, out var arg))
                    {
                        int maxA = 0, maxB = 0, maxC = 0;
                        foreach (var part in SplitSelectorList(arg))
                        {
                            var trimmed = part.Trim();
                            if (string.IsNullOrEmpty(trimmed)) continue;
                            var specificity = CalculateSelectorSpecificity(trimmed);
                            int a = specificity / 100;
                            int b = (specificity % 100) / 10;
                            int c2 = specificity % 10;
                            if (a > maxA || (a == maxA && b > maxB) || (a == maxA && b == maxB && c2 > maxC))
                            {
                                maxA = a;
                                maxB = b;
                                maxC = c2;
                            }
                        }
                        idCount += maxA;
                        classLikeCount += maxB;
                        typeCount += maxC;
                    }
                }
                else if (!pseudoElement && pseudoName.Equals("has", StringComparison.OrdinalIgnoreCase) && i < simple.Length && simple[i] == '(')
                {
                    if (TryReadFunctionArgument(simple, ref i, out var arg))
                    {
                        int maxA = 0, maxB = 0, maxC = 0;
                        foreach (var part in SplitSelectorList(arg))
                        {
                            var trimmed = part.Trim();
                            if (string.IsNullOrEmpty(trimmed)) continue;
                            var specificity = CalculateSelectorSpecificity(trimmed.TrimStart('>', '+', '~').Trim());
                            int a = specificity / 100;
                            int b = (specificity % 100) / 10;
                            int c2 = specificity % 10;
                            if (a > maxA || (a == maxA && b > maxB) || (a == maxA && b == maxB && c2 > maxC))
                            {
                                maxA = a;
                                maxB = b;
                                maxC = c2;
                            }
                        }
                        idCount += maxA;
                        classLikeCount += maxB;
                        typeCount += maxC;
                    }
                }
                else if (!pseudoElement && pseudoName.Equals("where", StringComparison.OrdinalIgnoreCase) && i < simple.Length && simple[i] == '(')
                {
                    TryReadFunctionArgument(simple, ref i, out _);
                }
                else if (!pseudoElement)
                {
                    classLikeCount++;
                    if (i < simple.Length && simple[i] == '(')
                        TryReadFunctionArgument(simple, ref i, out _);
                }
            }
            else
            {
                int start = i;
                while (i < simple.Length && simple[i] != '.' && simple[i] != '#' && simple[i] != '[' && simple[i] != ':') i++;
                var type = simple.Substring(start, i - start).Trim();
                if (!string.IsNullOrEmpty(type) && type != "*")
                    hasType = true;
            }
        }

        if (hasType) typeCount++;
        return (idCount, classLikeCount, typeCount);
    }

    private readonly struct StylePriority : IComparable<StylePriority>
    {
        public bool Important { get; }
        public int Specificity { get; }
        public int SourceOrder { get; }

        public StylePriority(bool important, int specificity, int sourceOrder)
        {
            Important = important;
            Specificity = specificity;
            SourceOrder = sourceOrder;
        }

        public int CompareTo(StylePriority other)
        {
            if (Important != other.Important) return Important ? 1 : -1;
            if (Specificity != other.Specificity) return Specificity.CompareTo(other.Specificity);
            return SourceOrder.CompareTo(other.SourceOrder);
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

    private static HtmlNode? GetNextElementSibling(HtmlNode node)
    {
        var parent = node.Parent;
        if (parent == null) return null;
        var siblings = parent.Children;
        var pos = siblings.IndexOf(node);
        if (pos < 0 || pos >= siblings.Count - 1) return null;
        for (int i = pos + 1; i < siblings.Count; i++)
        {
            var sib = siblings[i];
            if (sib.TagName == "#text") continue;
            return sib;
        }
        return null;
    }

    private static List<HtmlNode> GetFollowingElementSiblings(HtmlNode node)
    {
        var result = new List<HtmlNode>();
        var parent = node.Parent;
        if (parent == null) return result;
        var siblings = parent.Children;
        var pos = siblings.IndexOf(node);
        if (pos < 0 || pos >= siblings.Count - 1) return result;
        for (int i = pos + 1; i < siblings.Count; i++)
        {
            var sib = siblings[i];
            if (sib.TagName == "#text") continue;
            result.Add(sib);
        }
        return result;
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
            int parenDepth = 0;
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
                if (c == '(') { parenDepth++; i++; continue; }
                if (c == ')') { if (parenDepth > 0) parenDepth--; i++; continue; }

                if (bracketDepth == 0 && parenDepth == 0 && (char.IsWhiteSpace(c) || c == '>' || c == '+' || c == '~'))
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
        var attributes = new List<AttributeSelector>();
        var negations = new List<string>();

        int i = 0;
        while (i < sel.Length)
        {
            char c = sel[i];
            if (c == '#')
            {
                i++;
                int start = i;
                while (i < sel.Length && sel[i] != '.' && sel[i] != '#' && sel[i] != '[' && sel[i] != ':') i++;
                id = sel[start..i];
            }
            else if (c == '.')
            {
                i++;
                int start = i;
                while (i < sel.Length && sel[i] != '.' && sel[i] != '#' && sel[i] != '[' && sel[i] != ':') i++;
                classes.Add(sel[start..i]);
            }
            else if (c == '[')
            {
                i++;
                int start = i;
                while (i < sel.Length && sel[i] != ']') i++;
                var inside = sel[start..i];
                i++; // skip ']'
                ParseAttributeSelector(inside, attributes);
            }
            else if (c == ':')
            {
                i++;
                bool pseudoElement = i < sel.Length && sel[i] == ':';
                if (pseudoElement) return false;

                int nameStart = i;
                while (i < sel.Length && (char.IsLetterOrDigit(sel[i]) || sel[i] == '-' || sel[i] == '_')) i++;
                var pseudoName = sel.Substring(nameStart, i - nameStart);
                if (pseudoName.Equals("not", StringComparison.OrdinalIgnoreCase) && i < sel.Length && sel[i] == '(')
                {
                    if (!TryReadFunctionArgument(sel, ref i, out var arg)) return false;
                    foreach (var part in SplitSelectorList(arg))
                    {
                        var simplePart = part.Trim();
                        if (string.IsNullOrEmpty(simplePart)) return false;
                        negations.Add(simplePart);
                    }
                }
                else if ((pseudoName.Equals("is", StringComparison.OrdinalIgnoreCase) || pseudoName.Equals("where", StringComparison.OrdinalIgnoreCase)) && i < sel.Length && sel[i] == '(')
                {
                    if (!TryReadFunctionArgument(sel, ref i, out var arg)) return false;
                    bool any = false;
                    foreach (var part in SplitSelectorList(arg))
                    {
                        var simplePart = part.Trim();
                        if (string.IsNullOrEmpty(simplePart)) continue;
                        if (MatchesRelativeSelector(node, simplePart))
                        {
                            any = true;
                            break;
                        }
                    }
                    if (!any) return false;
                }
                else if (pseudoName.Equals("has", StringComparison.OrdinalIgnoreCase) && i < sel.Length && sel[i] == '(')
                {
                    if (!TryReadFunctionArgument(sel, ref i, out var arg)) return false;
                    bool any = false;
                    foreach (var part in SplitSelectorList(arg))
                    {
                        var relative = part.Trim();
                        if (string.IsNullOrEmpty(relative)) continue;
                        if (MatchesHasRelativeSelector(node, relative))
                        {
                            any = true;
                            break;
                        }
                    }
                    if (!any) return false;
                }
                else if (pseudoName.Equals("first-child", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsFirstChild(node)) return false;
                }
                else if (pseudoName.Equals("last-child", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsLastChild(node)) return false;
                }
                else if (pseudoName.Equals("only-child", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsOnlyChild(node)) return false;
                }
                else if (pseudoName.Equals("first-of-type", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsFirstOfType(node)) return false;
                }
                else if (pseudoName.Equals("last-of-type", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsLastOfType(node)) return false;
                }
                else if (pseudoName.Equals("only-of-type", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsOnlyOfType(node)) return false;
                }
                else if (pseudoName.Equals("nth-child", StringComparison.OrdinalIgnoreCase) && i < sel.Length && sel[i] == '(')
                {
                    if (!TryReadFunctionArgument(sel, ref i, out var arg)) return false;
                    if (!MatchesNthChild(node, arg)) return false;
                }
                else if (pseudoName.Equals("nth-of-type", StringComparison.OrdinalIgnoreCase) && i < sel.Length && sel[i] == '(')
                {
                    if (!TryReadFunctionArgument(sel, ref i, out var arg)) return false;
                    if (!MatchesNthOfType(node, arg)) return false;
                }
                else if (pseudoName.Equals("nth-last-child", StringComparison.OrdinalIgnoreCase) && i < sel.Length && sel[i] == '(')
                {
                    if (!TryReadFunctionArgument(sel, ref i, out var arg)) return false;
                    if (!MatchesNthLastChild(node, arg)) return false;
                }
                else if (pseudoName.Equals("nth-last-of-type", StringComparison.OrdinalIgnoreCase) && i < sel.Length && sel[i] == '(')
                {
                    if (!TryReadFunctionArgument(sel, ref i, out var arg)) return false;
                    if (!MatchesNthLastOfType(node, arg)) return false;
                }
                else if (pseudoName.Equals("empty", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsEmpty(node)) return false;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                // tag name
                int start = i;
                while (i < sel.Length && sel[i] != '.' && sel[i] != '#' && sel[i] != '[' && sel[i] != ':') i++;
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
            foreach (var attrSel in attributes)
            {
                var attr = node.GetAttribute(attrSel.Name);
                if (!MatchesAttribute(attr, attrSel))
                    return false;
            }
        }
        if (negations.Count > 0)
        {
            foreach (var neg in negations)
            {
                if (MatchesSimple(node, neg)) return false;
            }
        }

        return true;
    }

    private static bool MatchesRelativeSelector(HtmlNode node, string selector)
    {
        var tokens = TokenizeSelector(selector);
        if (tokens.Count == 0) return false;
        if (Combinators.Contains(tokens[^1])) return false;
        return MatchTokensFromRight(node, tokens, tokens.Count - 1);
    }

    private static bool MatchesHasRelativeSelector(HtmlNode scope, string selector)
    {
        var trimmed = selector.Trim();
        if (string.IsNullOrEmpty(trimmed)) return false;

        char lead = trimmed[0];
        if (lead == '>' || lead == '+' || lead == '~')
        {
            var rest = trimmed.Substring(1).Trim();
            if (string.IsNullOrEmpty(rest)) return false;
            bool searchDescendants = SelectorNeedsDescendantSearch(rest);
            var anchors = lead switch
            {
                '>' => scope.Children.Where(x => x.TagName != "#text").ToList(),
                '+' => GetNextElementSibling(scope) is HtmlNode n ? new List<HtmlNode> { n } : new List<HtmlNode>(),
                '~' => GetFollowingElementSiblings(scope),
                _ => new List<HtmlNode>()
            };
            foreach (var anchor in anchors)
            {
                if (!searchDescendants)
                {
                    if (MatchesRelativeSelector(anchor, rest)) return true;
                    continue;
                }

                foreach (var candidate in EnumerateSelfAndDescendants(anchor))
                {
                    if (MatchesRelativeSelector(candidate, rest)) return true;
                }
            }
            return false;
        }

        foreach (var descendant in EnumerateDescendants(scope))
        {
            if (MatchesRelativeSelector(descendant, trimmed)) return true;
        }
        return false;
    }

    private static bool IsFirstChild(HtmlNode node)
    {
        var siblings = GetElementSiblings(node);
        return siblings.Count > 0 && ReferenceEquals(siblings[0], node);
    }

    private static bool IsLastChild(HtmlNode node)
    {
        var siblings = GetElementSiblings(node);
        return siblings.Count > 0 && ReferenceEquals(siblings[^1], node);
    }

    private static bool IsOnlyChild(HtmlNode node)
    {
        var siblings = GetElementSiblings(node);
        return siblings.Count == 1 && ReferenceEquals(siblings[0], node);
    }

    private static bool IsFirstOfType(HtmlNode node)
    {
        var ofType = GetElementSiblings(node).Where(s => s.TagName.Equals(node.TagName, StringComparison.OrdinalIgnoreCase)).ToList();
        return ofType.Count > 0 && ReferenceEquals(ofType[0], node);
    }

    private static bool IsLastOfType(HtmlNode node)
    {
        var ofType = GetElementSiblings(node).Where(s => s.TagName.Equals(node.TagName, StringComparison.OrdinalIgnoreCase)).ToList();
        return ofType.Count > 0 && ReferenceEquals(ofType[^1], node);
    }

    private static bool IsOnlyOfType(HtmlNode node)
    {
        var ofType = GetElementSiblings(node).Where(s => s.TagName.Equals(node.TagName, StringComparison.OrdinalIgnoreCase)).ToList();
        return ofType.Count == 1 && ReferenceEquals(ofType[0], node);
    }

    private static bool MatchesNthChild(HtmlNode node, string expression)
    {
        var siblings = GetElementSiblings(node);
        int index = siblings.IndexOf(node) + 1;
        if (index <= 0) return false;
        return MatchesNthExpression(index, expression);
    }

    private static bool MatchesNthOfType(HtmlNode node, string expression)
    {
        var ofType = GetElementSiblings(node).Where(s => s.TagName.Equals(node.TagName, StringComparison.OrdinalIgnoreCase)).ToList();
        int index = ofType.IndexOf(node) + 1;
        if (index <= 0) return false;
        return MatchesNthExpression(index, expression);
    }

    private static bool MatchesNthLastChild(HtmlNode node, string expression)
    {
        var siblings = GetElementSiblings(node);
        int indexFromEnd = siblings.Count - siblings.IndexOf(node);
        if (indexFromEnd <= 0) return false;
        return MatchesNthExpression(indexFromEnd, expression);
    }

    private static bool MatchesNthLastOfType(HtmlNode node, string expression)
    {
        var ofType = GetElementSiblings(node).Where(s => s.TagName.Equals(node.TagName, StringComparison.OrdinalIgnoreCase)).ToList();
        int indexFromEnd = ofType.Count - ofType.IndexOf(node);
        if (indexFromEnd <= 0) return false;
        return MatchesNthExpression(indexFromEnd, expression);
    }

    private static bool IsEmpty(HtmlNode node)
    {
        if (node.Children.Count == 0) return true;
        foreach (var child in node.Children)
        {
            if (child.TagName == "#text")
            {
                if (!string.IsNullOrEmpty(child.Text))
                    return false;
            }
            else
            {
                return false;
            }
        }
        return true;
    }

    private static bool MatchesNthExpression(int index, string expression)
    {
        var expr = expression.Trim().ToLowerInvariant().Replace(" ", string.Empty);
        if (expr == "odd") return index % 2 == 1;
        if (expr == "even") return index % 2 == 0;
        if (int.TryParse(expr, out var exact)) return index == exact;

        int nPos = expr.IndexOf('n');
        if (nPos < 0) return false;

        var aPart = expr.Substring(0, nPos);
        var bPart = expr.Substring(nPos + 1);

        int a = aPart switch
        {
            "" or "+" => 1,
            "-" => -1,
            _ when int.TryParse(aPart, out var parsedA) => parsedA,
            _ => 0
        };
        if (a == 0) return false;

        int b = 0;
        if (!string.IsNullOrEmpty(bPart))
        {
            if (!int.TryParse(bPart, out b)) return false;
        }

        int diff = index - b;
        if (a > 0) return diff >= 0 && diff % a == 0;
        return diff <= 0 && diff % a == 0;
    }

    private static List<HtmlNode> GetElementSiblings(HtmlNode node)
    {
        if (node.Parent == null) return new List<HtmlNode> { node };
        return node.Parent.Children.Where(c => c.TagName != "#text").ToList();
    }

    private static bool SelectorNeedsDescendantSearch(string selector)
    {
        var tokens = TokenizeSelector(selector);
        return tokens.Any(t => Combinators.Contains(t));
    }

    private static IEnumerable<HtmlNode> EnumerateDescendants(HtmlNode node)
    {
        foreach (var child in node.Children)
        {
            if (child.TagName != "#text") yield return child;
            foreach (var inner in EnumerateDescendants(child)) yield return inner;
        }
    }

    private static IEnumerable<HtmlNode> EnumerateSelfAndDescendants(HtmlNode node)
    {
        if (node.TagName != "#text") yield return node;
        foreach (var d in EnumerateDescendants(node)) yield return d;
    }

    private static bool TryReadFunctionArgument(string source, ref int i, out string argument)
    {
        argument = string.Empty;
        if (i >= source.Length || source[i] != '(') return false;

        int start = ++i;
        int depth = 1;
        char quote = '\0';
        while (i < source.Length)
        {
            char c = source[i];
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
            if (c == '(') depth++;
            else if (c == ')')
            {
                depth--;
                if (depth == 0)
                {
                    argument = source.Substring(start, i - start);
                    i++;
                    return true;
                }
            }
            i++;
        }
        return false;
    }

    private static IEnumerable<string> SplitSelectorList(string value)
    {
        var parts = new List<string>();
        int start = 0;
        int bracketDepth = 0;
        int parenDepth = 0;
        char quote = '\0';
        for (int i = 0; i < value.Length; i++)
        {
            char c = value[i];
            if (quote != '\0')
            {
                if (c == quote) quote = '\0';
                continue;
            }
            if (c == '"' || c == '\'')
            {
                quote = c;
                continue;
            }
            if (c == '[') { bracketDepth++; continue; }
            if (c == ']') { if (bracketDepth > 0) bracketDepth--; continue; }
            if (c == '(') { parenDepth++; continue; }
            if (c == ')') { if (parenDepth > 0) parenDepth--; continue; }
            if (c == ',' && bracketDepth == 0 && parenDepth == 0)
            {
                parts.Add(value.Substring(start, i - start));
                start = i + 1;
            }
        }
        parts.Add(value.Substring(start));
        return parts;
    }

    private static void ParseAttributeSelector(string inside, List<AttributeSelector> attributes)
    {
        string raw = inside.Trim();
        if (string.IsNullOrEmpty(raw)) return;
        string[] operators = { "~=", "|=", "^=", "$=", "*=", "=" };
        foreach (var op in operators)
        {
            int idx = raw.IndexOf(op, StringComparison.Ordinal);
            if (idx > 0)
            {
                var name = raw.Substring(0, idx).Trim();
                var value = raw.Substring(idx + op.Length).Trim().Trim('"', '\'');
                attributes.Add(new AttributeSelector(name, op, value));
                return;
            }
        }
        attributes.Add(new AttributeSelector(raw, null, null));
    }

    private static bool MatchesAttribute(string? attr, AttributeSelector sel)
    {
        if (attr == null) return false;
        if (sel.Operator == null) return true;
        var value = sel.Value ?? string.Empty;
        return sel.Operator switch
        {
            "=" => attr.Equals(value, StringComparison.OrdinalIgnoreCase),
            "^=" => attr.StartsWith(value, StringComparison.OrdinalIgnoreCase),
            "$=" => attr.EndsWith(value, StringComparison.OrdinalIgnoreCase),
            "*=" => attr.Contains(value, StringComparison.OrdinalIgnoreCase),
            "~=" => attr.Split(' ', StringSplitOptions.RemoveEmptyEntries).Any(x => x.Equals(value, StringComparison.OrdinalIgnoreCase)),
            "|=" => attr.Equals(value, StringComparison.OrdinalIgnoreCase) || attr.StartsWith(value + "-", StringComparison.OrdinalIgnoreCase),
            _ => false
        };
    }

    private readonly struct AttributeSelector
    {
        public string Name { get; }
        public string? Operator { get; }
        public string? Value { get; }

        public AttributeSelector(string name, string? @operator, string? value)
        {
            Name = name;
            Operator = @operator;
            Value = value;
        }
    }
}
