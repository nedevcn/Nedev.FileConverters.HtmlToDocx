using System;
using System.Collections.Generic;

namespace Nedev.FileConverters.HtmlToDocx.Core.Css;

public sealed class CssParser
{
    private static readonly HashSet<string> SupportsLikelyProperties = new(StringComparer.OrdinalIgnoreCase)
    {
        "color", "background-color", "font-family", "font-size", "font-weight", "font-style", "text-decoration",
        "text-align", "line-height",
        "margin", "margin-top", "margin-right", "margin-bottom", "margin-left",
        "padding", "padding-top", "padding-right", "padding-bottom", "padding-left",
        "border", "border-top", "border-right", "border-bottom", "border-left"
    };

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

        int start = 0;
        int sourceOrder = 0;

        while (start < style.Length)
        {
            SkipWhitespaceAndComments(style, ref start);
            if (start >= style.Length) break;

            int colonPos = FindTopLevelChar(style, ':', start);
            if (colonPos < 0) break;

            int semicolonPos = FindTopLevelChar(style, ';', colonPos + 1);
            var name = style.Substring(start, colonPos - start).Trim().ToLowerInvariant();
            string value;
            if (semicolonPos < 0)
            {
                value = style.Substring(colonPos + 1).Trim();
                start = style.Length;
            }
            else
            {
                value = style.Substring(colonPos + 1, semicolonPos - colonPos - 1).Trim();
                start = semicolonPos + 1;
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

        int pos = 0;
        while (pos < css.Length)
        {
            SkipWhitespaceAndComments(css, ref pos);
            if (pos >= css.Length) break;

            if (css[pos] == '@')
            {
                if (!TryConsumeAtRule(css, ref pos, rules))
                    break;
                continue;
            }

            int braceOpen = FindTopLevelChar(css, '{', pos);
            if (braceOpen < 0) break;
            var rawSelector = css.Substring(pos, braceOpen - pos).Trim();
            if (!TryReadBlock(css, braceOpen, out var body, out var afterBlock))
                break;
            var selectors = SplitSelectors(rawSelector);
            foreach (var sel in selectors)
            {
                rules.Add(new CssRule(sel.Trim(), ParseDeclarations(body)));
            }

            pos = afterBlock;
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

    private static List<string> SplitSelectors(string rawSelector)
    {
        var selectors = new List<string>();
        if (string.IsNullOrWhiteSpace(rawSelector)) return selectors;

        int start = 0;
        int bracketDepth = 0;
        int parenDepth = 0;
        char quote = '\0';

        for (int i = 0; i < rawSelector.Length; i++)
        {
            char c = rawSelector[i];
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
                var part = rawSelector.Substring(start, i - start).Trim();
                if (!string.IsNullOrEmpty(part)) selectors.Add(part);
                start = i + 1;
            }
        }

        var last = rawSelector.Substring(start).Trim();
        if (!string.IsNullOrEmpty(last)) selectors.Add(last);
        return selectors;
    }

    private static void SkipWhitespaceAndComments(string css, ref int pos)
    {
        while (pos < css.Length)
        {
            if (char.IsWhiteSpace(css[pos]))
            {
                pos++;
                continue;
            }

            if (pos + 1 < css.Length && css[pos] == '/' && css[pos + 1] == '*')
            {
                int end = css.IndexOf("*/", pos + 2, StringComparison.Ordinal);
                if (end < 0)
                {
                    pos = css.Length;
                    return;
                }
                pos = end + 2;
                continue;
            }

            break;
        }
    }

    private static bool TryConsumeAtRule(string css, ref int pos, List<CssRule> rules)
    {
        int start = pos;
        pos++;
        while (pos < css.Length && (char.IsLetterOrDigit(css[pos]) || css[pos] == '-')) pos++;
        var atName = css.Substring(start + 1, pos - start - 1);

        SkipWhitespaceAndComments(css, ref pos);
        int headerStart = pos;
        int bracePos = FindTopLevelChar(css, '{', pos);
        int semiPos = FindTopLevelChar(css, ';', pos);

        if (semiPos >= 0 && (bracePos < 0 || semiPos < bracePos))
        {
            pos = semiPos + 1;
            return true;
        }

        if (bracePos < 0) return false;
        var prelude = css.Substring(headerStart, bracePos - headerStart).Trim();
        if (!TryReadBlock(css, bracePos, out var blockBody, out var afterBlock))
            return false;

        if (atName.Equals("media", StringComparison.OrdinalIgnoreCase) && ShouldApplyMedia(prelude))
        {
            rules.AddRange(ParseStylesheet(blockBody));
        }
        else if (atName.Equals("supports", StringComparison.OrdinalIgnoreCase) && ShouldApplySupports(prelude))
        {
            rules.AddRange(ParseStylesheet(blockBody));
        }

        pos = afterBlock;
        return true;
    }

    private static bool ShouldApplyMedia(string query)
    {
        if (string.IsNullOrWhiteSpace(query)) return true;
        foreach (var part in query.Split(',', StringSplitOptions.RemoveEmptyEntries))
        {
            var q = part.Trim().ToLowerInvariant();
            if (q.Contains("print", StringComparison.Ordinal)) return true;
            if (q.Contains("all", StringComparison.Ordinal)) return true;
        }
        return false;
    }

    private static bool ShouldApplySupports(string condition)
    {
        if (string.IsNullOrWhiteSpace(condition)) return false;
        return EvaluateSupportsExpression(condition.Trim());
    }

    private static bool EvaluateSupportsExpression(string expression)
    {
        var expr = StripOuterParentheses(expression.Trim());
        if (expr.Length == 0) return false;

        if (expr.StartsWith("not ", StringComparison.OrdinalIgnoreCase))
            return !EvaluateSupportsExpression(expr.Substring(4).Trim());

        var orParts = SplitTopLevelByKeyword(expr, "or");
        if (orParts.Count > 1) return orParts.Any(EvaluateSupportsExpression);

        var andParts = SplitTopLevelByKeyword(expr, "and");
        if (andParts.Count > 1) return andParts.All(EvaluateSupportsExpression);

        return EvaluateSupportsTerm(expr);
    }

    private static bool EvaluateSupportsTerm(string expression)
    {
        var term = StripOuterParentheses(expression.Trim());
        if (term.Length == 0) return false;

        int colon = FindTopLevelChar(term, ':', 0);
        if (colon <= 0) return false;

        var property = term.Substring(0, colon).Trim();
        var value = term.Substring(colon + 1).Trim();
        if (property.Length == 0 || value.Length == 0) return false;

        if (property.StartsWith("--", StringComparison.Ordinal)) return true;
        return SupportsLikelyProperties.Contains(property);
    }

    private static List<string> SplitTopLevelByKeyword(string input, string keyword)
    {
        var parts = new List<string>();
        int start = 0;
        int bracketDepth = 0;
        int parenDepth = 0;
        char quote = '\0';
        for (int i = 0; i < input.Length; i++)
        {
            char c = input[i];
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
            if (bracketDepth != 0 || parenDepth != 0) continue;

            if (IsKeywordAt(input, i, keyword))
            {
                var part = input.Substring(start, i - start).Trim();
                if (part.Length > 0) parts.Add(part);
                i += keyword.Length - 1;
                start = i + 1;
            }
        }

        var last = input.Substring(start).Trim();
        if (last.Length > 0) parts.Add(last);
        return parts;
    }

    private static bool IsKeywordAt(string input, int index, string keyword)
    {
        if (index + keyword.Length > input.Length) return false;
        if (!input.AsSpan(index, keyword.Length).Equals(keyword.AsSpan(), StringComparison.OrdinalIgnoreCase))
            return false;
        var beforeOk = index == 0 || char.IsWhiteSpace(input[index - 1]);
        var afterIndex = index + keyword.Length;
        var afterOk = afterIndex >= input.Length || char.IsWhiteSpace(input[afterIndex]);
        return beforeOk && afterOk;
    }

    private static string StripOuterParentheses(string input)
    {
        var s = input.Trim();
        while (s.Length >= 2 && s[0] == '(' && s[^1] == ')')
        {
            int depth = 0;
            bool wrapsAll = true;
            char quote = '\0';
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
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
                if (c == '(') depth++;
                else if (c == ')')
                {
                    depth--;
                    if (depth == 0 && i != s.Length - 1)
                    {
                        wrapsAll = false;
                        break;
                    }
                }
            }
            if (!wrapsAll) break;
            s = s.Substring(1, s.Length - 2).Trim();
        }
        return s;
    }

    private static bool TryReadBlock(string css, int braceOpenIndex, out string body, out int nextPos)
    {
        body = string.Empty;
        nextPos = braceOpenIndex;
        if (braceOpenIndex < 0 || braceOpenIndex >= css.Length || css[braceOpenIndex] != '{') return false;

        int depth = 0;
        char quote = '\0';
        for (int i = braceOpenIndex; i < css.Length; i++)
        {
            char c = css[i];
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

            if (c == '/' && i + 1 < css.Length && css[i + 1] == '*')
            {
                int end = css.IndexOf("*/", i + 2, StringComparison.Ordinal);
                if (end < 0) return false;
                i = end + 1;
                continue;
            }

            if (c == '{') depth++;
            else if (c == '}')
            {
                depth--;
                if (depth == 0)
                {
                    body = css.Substring(braceOpenIndex + 1, i - braceOpenIndex - 1);
                    nextPos = i + 1;
                    return true;
                }
            }
        }
        return false;
    }

    private static int FindTopLevelChar(string css, char target, int start)
    {
        int bracketDepth = 0;
        int parenDepth = 0;
        char quote = '\0';
        for (int i = start; i < css.Length; i++)
        {
            char c = css[i];
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
            if (c == '/' && i + 1 < css.Length && css[i + 1] == '*')
            {
                int end = css.IndexOf("*/", i + 2, StringComparison.Ordinal);
                if (end < 0) return -1;
                i = end + 1;
                continue;
            }
            if (c == '[') { bracketDepth++; continue; }
            if (c == ']') { if (bracketDepth > 0) bracketDepth--; continue; }
            if (c == '(') { parenDepth++; continue; }
            if (c == ')') { if (parenDepth > 0) parenDepth--; continue; }
            if (bracketDepth == 0 && parenDepth == 0 && c == target) return i;
        }
        return -1;
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
