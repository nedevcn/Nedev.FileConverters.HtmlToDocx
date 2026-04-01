using System;
using System.Collections.Generic;
using System.Globalization;
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
            if (InheritableProperties.Contains(kvp.Key) || kvp.Key.StartsWith("--", StringComparison.Ordinal))
            {
                computed[kvp.Key] = kvp.Value;
                priority[kvp.Key] = new StylePriority(important: false, specificity: Specificity.None, sourceOrder: -1);
            }
        }

        int sourceOrder = 0;
        foreach (var rule in stylesheet)
        {
            if (MatchesSelector(node, rule.Selector))
            {
                var specificity = CalculateSelectorSpecificity(rule.Selector);
                foreach (var decl in rule.Declarations)
                {
                    int declarationOrder = unchecked(sourceOrder * 1000 + decl.SourceOrder);
                    ApplyDeclaration(node, computed, priority, decl.Property, decl.Value, decl.Important, specificity, declarationOrder);
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
                    node,
                    computed,
                    priority,
                    decl.Property,
                    decl.Value,
                    decl.Important,
                    specificity: Specificity.InlineStyle,
                    sourceOrder: int.MaxValue - inlineDecls.Count + decl.SourceOrder
                );
            }
        }

        ResolveComputedValues(computed, parentStyle);

        node.ComputedStyle = computed;

        foreach (var child in node.Children)
        {
            ResolveNodeStyle(child, stylesheet, computed);
        }
    }

    private static void ApplyDeclaration(
        HtmlNode node,
        Dictionary<string, string> computed,
        Dictionary<string, StylePriority> priority,
        string property,
        string value,
        bool important,
        Specificity specificity,
        int sourceOrder)
    {
        var incoming = new StylePriority(important, specificity, sourceOrder);
        foreach (var (targetProperty, targetValue) in ExpandLogicalProperty(node, property, value))
        {
            if (!priority.TryGetValue(targetProperty, out var existing) || incoming.CompareTo(existing) >= 0)
            {
                computed[targetProperty] = targetValue;
                priority[targetProperty] = incoming;
            }
        }
    }

    private static IEnumerable<(string property, string value)> ExpandLogicalProperty(HtmlNode node, string property, string value)
    {
        var dir = ResolveDirection(node);
        bool rtl = dir.Equals("rtl", StringComparison.OrdinalIgnoreCase);
        switch (property.ToLowerInvariant())
        {
            case "margin-inline-start":
                yield return (rtl ? "margin-right" : "margin-left", value);
                yield break;
            case "margin-inline-end":
                yield return (rtl ? "margin-left" : "margin-right", value);
                yield break;
            case "padding-inline-start":
                yield return (rtl ? "padding-right" : "padding-left", value);
                yield break;
            case "padding-inline-end":
                yield return (rtl ? "padding-left" : "padding-right", value);
                yield break;
            case "margin-block-start":
                yield return ("margin-top", value);
                yield break;
            case "margin-block-end":
                yield return ("margin-bottom", value);
                yield break;
            case "padding-block-start":
                yield return ("padding-top", value);
                yield break;
            case "padding-block-end":
                yield return ("padding-bottom", value);
                yield break;
            case "margin-inline":
            {
                var parts = SplitTopLevelWhitespace(value);
                if (parts.Count == 1)
                {
                    yield return (rtl ? "margin-right" : "margin-left", parts[0]);
                    yield return (rtl ? "margin-left" : "margin-right", parts[0]);
                }
                else if (parts.Count >= 2)
                {
                    yield return (rtl ? "margin-right" : "margin-left", parts[0]);
                    yield return (rtl ? "margin-left" : "margin-right", parts[1]);
                }
                else
                {
                    yield return (property, value);
                }
                yield break;
            }
            case "padding-inline":
            {
                var parts = SplitTopLevelWhitespace(value);
                if (parts.Count == 1)
                {
                    yield return (rtl ? "padding-right" : "padding-left", parts[0]);
                    yield return (rtl ? "padding-left" : "padding-right", parts[0]);
                }
                else if (parts.Count >= 2)
                {
                    yield return (rtl ? "padding-right" : "padding-left", parts[0]);
                    yield return (rtl ? "padding-left" : "padding-right", parts[1]);
                }
                else
                {
                    yield return (property, value);
                }
                yield break;
            }
            case "margin-block":
            {
                var parts = SplitTopLevelWhitespace(value);
                if (parts.Count == 1)
                {
                    yield return ("margin-top", parts[0]);
                    yield return ("margin-bottom", parts[0]);
                }
                else if (parts.Count >= 2)
                {
                    yield return ("margin-top", parts[0]);
                    yield return ("margin-bottom", parts[1]);
                }
                else
                {
                    yield return (property, value);
                }
                yield break;
            }
            case "padding-block":
            {
                var parts = SplitTopLevelWhitespace(value);
                if (parts.Count == 1)
                {
                    yield return ("padding-top", parts[0]);
                    yield return ("padding-bottom", parts[0]);
                }
                else if (parts.Count >= 2)
                {
                    yield return ("padding-top", parts[0]);
                    yield return ("padding-bottom", parts[1]);
                }
                else
                {
                    yield return (property, value);
                }
                yield break;
            }
            case "font":
            {
                if (TryExpandFontShorthand(value, out var expanded))
                {
                    foreach (var item in expanded)
                        yield return item;
                }
                else
                {
                    yield return (property, value);
                }
                yield break;
            }
            default:
                yield return (property, value);
                yield break;
        }
    }

    private static string ResolveDirection(HtmlNode node)
    {
        HtmlNode? current = node;
        while (current != null)
        {
            var dir = current.GetAttribute("dir");
            if (!string.IsNullOrWhiteSpace(dir))
                return dir.Trim();
            current = current.Parent;
        }
        return "ltr";
    }

    private static List<string> SplitTopLevelWhitespace(string input)
    {
        var result = new List<string>();
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
            if (!char.IsWhiteSpace(c) || bracketDepth != 0 || parenDepth != 0) continue;

            if (i > start)
                result.Add(input.Substring(start, i - start).Trim());
            while (i + 1 < input.Length && char.IsWhiteSpace(input[i + 1])) i++;
            start = i + 1;
        }
        if (start < input.Length)
        {
            var tail = input.Substring(start).Trim();
            if (tail.Length > 0) result.Add(tail);
        }
        return result;
    }

    private static bool TryExpandFontShorthand(string value, out List<(string property, string value)> expanded)
    {
        expanded = new List<(string property, string value)>();
        var tokens = SplitTopLevelWhitespace(value);
        if (tokens.Count < 2) return false;

        int sizeIndex = -1;
        string sizeToken = string.Empty;
        for (int i = 0; i < tokens.Count; i++)
        {
            if (TryParseFontSizeToken(tokens[i], out var size, out var lineHeight))
            {
                sizeIndex = i;
                sizeToken = size;
                if (!string.IsNullOrWhiteSpace(lineHeight))
                    expanded.Add(("line-height", lineHeight!));
                break;
            }
        }

        if (sizeIndex < 0 || sizeIndex >= tokens.Count - 1) return false;

        string? style = null;
        string? weight = null;
        for (int i = 0; i < sizeIndex; i++)
        {
            var t = tokens[i].ToLowerInvariant();
            if (style == null && (t == "italic" || t == "oblique" || t == "normal"))
            {
                style = tokens[i];
                continue;
            }
            if (weight == null && (t == "bold" || t == "bolder" || t == "lighter" || IsNumericFontWeight(t)))
            {
                weight = tokens[i];
                continue;
            }
        }

        var family = string.Join(" ", tokens.Skip(sizeIndex + 1)).Trim();
        if (string.IsNullOrWhiteSpace(family)) return false;

        expanded.Add(("font-size", sizeToken));
        expanded.Add(("font-family", family));
        if (style != null) expanded.Add(("font-style", style));
        if (weight != null) expanded.Add(("font-weight", weight));
        return true;
    }

    private static bool TryParseFontSizeToken(string token, out string size, out string? lineHeight)
    {
        size = string.Empty;
        lineHeight = null;
        var parts = token.Split('/', 2, StringSplitOptions.TrimEntries);
        var maybeSize = parts[0].Trim().ToLowerInvariant();
        if (!IsValidFontSize(maybeSize)) return false;
        size = parts[0].Trim();
        if (parts.Length == 2 && parts[1].Length > 0)
            lineHeight = parts[1].Trim();
        return true;
    }

    private static bool IsValidFontSize(string token)
    {
        if (token.EndsWith("px", StringComparison.Ordinal) ||
            token.EndsWith("pt", StringComparison.Ordinal) ||
            token.EndsWith("em", StringComparison.Ordinal) ||
            token.EndsWith("rem", StringComparison.Ordinal) ||
            token.EndsWith("%", StringComparison.Ordinal))
            return true;

        return token == "xx-small" || token == "x-small" || token == "small" || token == "medium" ||
               token == "large" || token == "x-large" || token == "xx-large" || token == "smaller" || token == "larger";
    }

    private static bool IsNumericFontWeight(string token)
    {
        if (!int.TryParse(token, out var n)) return false;
        return n >= 100 && n <= 900 && n % 100 == 0;
    }

    private static void ResolveComputedValues(Dictionary<string, string> computed, Dictionary<string, string> parentStyle)
    {
        var keys = computed.Keys.Where(k => !k.StartsWith("--", StringComparison.Ordinal)).ToList();
        foreach (var key in keys)
        {
            if (!computed.TryGetValue(key, out var raw)) continue;
            if (!ResolveCssWideKeywordValue(key, raw, parentStyle, out var normalized))
            {
                computed.Remove(key);
                continue;
            }
            if (normalized == null)
            {
                computed.Remove(key);
                continue;
            }

            if (!TryResolveCssVariables(normalized, computed, new HashSet<string>(StringComparer.OrdinalIgnoreCase), out var resolved)) continue;
            if (TryResolveCalcFunctions(resolved, out var calcResolved))
                resolved = calcResolved;
            if (TryResolveMinMaxClampFunctions(resolved, out var minMaxClampResolved))
                resolved = minMaxClampResolved;
            computed[key] = resolved;
        }

        ResolveCurrentColorReferences(computed);
    }

    private static bool ResolveCssWideKeywordValue(string property, string rawValue, Dictionary<string, string> parentStyle, out string? normalized)
    {
        normalized = rawValue;
        var keyword = rawValue.Trim().ToLowerInvariant();
        if (keyword == "inherit")
        {
            if (parentStyle.TryGetValue(property, out var inherited))
            {
                normalized = inherited;
                return true;
            }
            normalized = null;
            return true;
        }

        if (keyword == "unset")
        {
            if (InheritableProperties.Contains(property) && parentStyle.TryGetValue(property, out var inherited))
            {
                normalized = inherited;
                return true;
            }
            normalized = null;
            return true;
        }

        if (keyword == "initial" || keyword == "revert")
        {
            normalized = null;
            return true;
        }

        return true;
    }

    private static void ResolveCurrentColorReferences(Dictionary<string, string> computed)
    {
        if (!computed.TryGetValue("color", out var currentColor) || string.IsNullOrWhiteSpace(currentColor)) return;
        var keys = computed.Keys.Where(k => !k.StartsWith("--", StringComparison.Ordinal) && !k.Equals("color", StringComparison.OrdinalIgnoreCase)).ToList();
        foreach (var key in keys)
        {
            if (!computed.TryGetValue(key, out var value)) continue;
            if (string.IsNullOrEmpty(value) || value.IndexOf("currentColor", StringComparison.OrdinalIgnoreCase) < 0) continue;
            computed[key] = ReplaceIgnoreCase(value, "currentColor", currentColor);
        }
    }

    private static string ReplaceIgnoreCase(string input, string oldValue, string newValue)
    {
        if (string.IsNullOrEmpty(input) || string.IsNullOrEmpty(oldValue)) return input;
        int pos = 0;
        var sb = new System.Text.StringBuilder(input.Length);
        while (pos < input.Length)
        {
            int idx = input.IndexOf(oldValue, pos, StringComparison.OrdinalIgnoreCase);
            if (idx < 0)
            {
                sb.Append(input, pos, input.Length - pos);
                break;
            }
            sb.Append(input, pos, idx - pos);
            sb.Append(newValue);
            pos = idx + oldValue.Length;
        }
        return sb.ToString();
    }

    private static bool TryResolveCalcFunctions(string value, out string resolved)
    {
        resolved = value;
        if (string.IsNullOrEmpty(value) || !value.Contains("calc(", StringComparison.OrdinalIgnoreCase))
            return true;

        var sb = new System.Text.StringBuilder(value.Length);
        int i = 0;
        while (i < value.Length)
        {
            if (i + 5 <= value.Length && value.AsSpan(i, 5).Equals("calc(".AsSpan(), StringComparison.OrdinalIgnoreCase))
            {
                int p = i + 4;
                if (!TryReadFunctionArgument(value, ref p, out var arg))
                {
                    sb.Append(value[i]);
                    i++;
                    continue;
                }

                if (TryEvaluateCalcExpression(arg, out var evaluated))
                    sb.Append(evaluated);
                else
                    sb.Append(value.Substring(i, p - i));

                i = p;
                continue;
            }

            sb.Append(value[i]);
            i++;
        }

        resolved = sb.ToString();
        return true;
    }

    private static bool TryResolveMinMaxClampFunctions(string value, out string resolved)
    {
        resolved = value;
        if (string.IsNullOrEmpty(value)) return true;
        if (value.IndexOf("min(", StringComparison.OrdinalIgnoreCase) < 0 &&
            value.IndexOf("max(", StringComparison.OrdinalIgnoreCase) < 0 &&
            value.IndexOf("clamp(", StringComparison.OrdinalIgnoreCase) < 0)
            return true;

        var sb = new System.Text.StringBuilder(value.Length);
        int i = 0;
        while (i < value.Length)
        {
            if (TryMatchFunctionName(value, i, "min", out var matchedLen) ||
                TryMatchFunctionName(value, i, "max", out matchedLen) ||
                TryMatchFunctionName(value, i, "clamp", out matchedLen))
            {
                var fn = value.Substring(i, matchedLen).ToLowerInvariant();
                int p = i + matchedLen;
                if (!TryReadFunctionArgument(value, ref p, out var arg))
                {
                    sb.Append(value[i]);
                    i++;
                    continue;
                }

                if (TryEvaluateMinMaxClamp(fn, arg, out var evaluated))
                    sb.Append(evaluated);
                else
                    sb.Append(value.Substring(i, p - i));

                i = p;
                continue;
            }

            sb.Append(value[i]);
            i++;
        }

        resolved = sb.ToString();
        return true;
    }

    private static bool TryMatchFunctionName(string input, int start, string name, out int length)
    {
        length = 0;
        if (start + name.Length + 1 > input.Length) return false;
        if (!input.AsSpan(start, name.Length).Equals(name.AsSpan(), StringComparison.OrdinalIgnoreCase)) return false;
        if (input[start + name.Length] != '(') return false;
        length = name.Length;
        return true;
    }

    private static bool TryEvaluateMinMaxClamp(string functionName, string argumentList, out string result)
    {
        result = string.Empty;
        var args = SplitFunctionArguments(argumentList);
        if (functionName == "clamp")
        {
            if (args.Count != 3) return false;
            if (!TryParseCssDimension(args[0], out var minNum, out var minUnit)) return false;
            if (!TryParseCssDimension(args[1], out var valNum, out var valUnit)) return false;
            if (!TryParseCssDimension(args[2], out var maxNum, out var maxUnit)) return false;
            if (!TryNormalizeUnitsForMath(ref minNum, ref minUnit, ref valNum, ref valUnit)) return false;
            if (!TryNormalizeUnitsForMath(ref valNum, ref valUnit, ref maxNum, ref maxUnit)) return false;
            var clamped = Math.Min(Math.Max(valNum, minNum), maxNum);
            result = clamped.ToString("0.#######", CultureInfo.InvariantCulture) + valUnit;
            return true;
        }

        if (args.Count < 2) return false;
        if (!TryParseCssDimension(args[0], out var accNum, out var accUnit)) return false;
        for (int i = 1; i < args.Count; i++)
        {
            if (!TryParseCssDimension(args[i], out var num, out var unit)) return false;
            if (!TryNormalizeUnitsForMath(ref accNum, ref accUnit, ref num, ref unit)) return false;
            accNum = functionName == "min" ? Math.Min(accNum, num) : Math.Max(accNum, num);
        }
        result = accNum.ToString("0.#######", CultureInfo.InvariantCulture) + accUnit;
        return true;
    }

    private static bool TryNormalizeUnitsForMath(ref decimal leftNum, ref string leftUnit, ref decimal rightNum, ref string rightUnit)
    {
        if (string.Equals(leftUnit, rightUnit, StringComparison.OrdinalIgnoreCase)) return true;
        if (string.IsNullOrEmpty(leftUnit) && leftNum == 0m)
        {
            leftUnit = rightUnit;
            return true;
        }
        if (string.IsNullOrEmpty(rightUnit) && rightNum == 0m)
        {
            rightUnit = leftUnit;
            return true;
        }
        return false;
    }

    private static List<string> SplitFunctionArguments(string args)
    {
        var result = new List<string>();
        int start = 0;
        int bracketDepth = 0;
        int parenDepth = 0;
        char quote = '\0';
        for (int i = 0; i < args.Length; i++)
        {
            char c = args[i];
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
                var part = args.Substring(start, i - start).Trim();
                if (part.Length > 0) result.Add(part);
                start = i + 1;
            }
        }
        var last = args.Substring(start).Trim();
        if (last.Length > 0) result.Add(last);
        return result;
    }

    private static bool TryEvaluateCalcExpression(string expression, out string result)
    {
        result = string.Empty;
        if (!SplitCalcTerms(expression, out var terms) || terms.Count == 0)
            return false;

        decimal total = 0m;
        string unit = string.Empty;
        foreach (var term in terms)
        {
            if (!TryParseCssDimension(term.term, out var num, out var termUnit))
                return false;

            if (string.IsNullOrEmpty(unit))
            {
                if (!string.IsNullOrEmpty(termUnit))
                    unit = termUnit;
            }
            else if (!string.IsNullOrEmpty(termUnit) && !termUnit.Equals(unit, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (string.IsNullOrEmpty(termUnit) && num != 0m && !string.IsNullOrEmpty(unit))
                return false;

            total += term.sign * num;
        }

        result = total.ToString("0.#######", CultureInfo.InvariantCulture) + unit;
        return true;
    }

    private static bool SplitCalcTerms(string expression, out List<(int sign, string term)> terms)
    {
        terms = new List<(int sign, string term)>();
        var expr = expression.Trim();
        if (expr.Length == 0) return false;

        int start = 0;
        int sign = 1;
        int parenDepth = 0;
        char quote = '\0';

        for (int i = 0; i < expr.Length; i++)
        {
            var c = expr[i];
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
            if (c == '(') { parenDepth++; continue; }
            if (c == ')') { if (parenDepth > 0) parenDepth--; continue; }

            if (parenDepth == 0 && (c == '+' || c == '-'))
            {
                var between = expr.Substring(start, i - start).Trim();
                if (between.Length == 0)
                {
                    sign = c == '-' ? -1 : 1;
                    start = i + 1;
                    continue;
                }

                terms.Add((sign, between));
                sign = c == '-' ? -1 : 1;
                start = i + 1;
            }
        }

        var last = expr.Substring(start).Trim();
        if (last.Length == 0) return false;
        terms.Add((sign, last));
        return true;
    }

    private static bool TryParseCssDimension(string rawTerm, out decimal number, out string unit)
    {
        number = 0m;
        unit = string.Empty;
        var term = rawTerm.Trim();
        if (term.StartsWith("calc(", StringComparison.OrdinalIgnoreCase))
        {
            if (!TryResolveCalcFunctions(term, out var collapsed))
                return false;
            term = collapsed.Trim();
        }

        if (term.StartsWith("(") && term.EndsWith(")", StringComparison.Ordinal) && term.Length > 2)
            term = term.Substring(1, term.Length - 2).Trim();

        int i = 0;
        if (i < term.Length && (term[i] == '+' || term[i] == '-')) i++;
        bool seenDigit = false;
        while (i < term.Length && (char.IsDigit(term[i]) || term[i] == '.'))
        {
            if (char.IsDigit(term[i])) seenDigit = true;
            i++;
        }

        if (!seenDigit) return false;
        var numPart = term.Substring(0, i).Trim();
        if (!decimal.TryParse(numPart, NumberStyles.Number, CultureInfo.InvariantCulture, out number))
            return false;

        unit = term.Substring(i).Trim().ToLowerInvariant();
        return true;
    }

    private static bool TryResolveCssVariables(
        string value,
        Dictionary<string, string> computed,
        HashSet<string> varStack,
        out string resolved)
    {
        resolved = value;
        if (string.IsNullOrEmpty(value) || !value.Contains("var(", StringComparison.Ordinal))
            return true;

        var sb = new System.Text.StringBuilder(value.Length);
        int i = 0;
        while (i < value.Length)
        {
            if (i + 4 <= value.Length && value.AsSpan(i, 4).SequenceEqual("var(".AsSpan()))
            {
                int p = i + 3;
                if (!TryReadFunctionArgument(value, ref p, out var arg))
                    return false;

                SplitVarFunctionArgument(arg, out var varName, out var fallback);
                if (string.IsNullOrEmpty(varName) || !varName.StartsWith("--", StringComparison.Ordinal))
                    return false;

                string replacement;
                if (computed.TryGetValue(varName, out var varValue))
                {
                    if (varStack.Contains(varName))
                    {
                        if (fallback == null) return false;
                        if (!TryResolveCssVariables(fallback, computed, varStack, out replacement)) return false;
                    }
                    else
                    {
                        varStack.Add(varName);
                        var ok = TryResolveCssVariables(varValue, computed, varStack, out replacement);
                        varStack.Remove(varName);
                        if (!ok)
                        {
                            if (fallback == null) return false;
                            if (!TryResolveCssVariables(fallback, computed, varStack, out replacement)) return false;
                        }
                    }
                }
                else
                {
                    if (fallback == null) return false;
                    if (!TryResolveCssVariables(fallback, computed, varStack, out replacement)) return false;
                }

                sb.Append(replacement);
                i = p;
                continue;
            }

            sb.Append(value[i]);
            i++;
        }

        resolved = sb.ToString();
        return true;
    }

    private static void SplitVarFunctionArgument(string arg, out string varName, out string? fallback)
    {
        varName = arg.Trim();
        fallback = null;
        int bracketDepth = 0;
        int parenDepth = 0;
        char quote = '\0';
        for (int i = 0; i < arg.Length; i++)
        {
            char c = arg[i];
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
                varName = arg.Substring(0, i).Trim();
                fallback = arg.Substring(i + 1).Trim();
                return;
            }
        }
    }

    private static Specificity CalculateSelectorSpecificity(string selector)
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

        return Specificity.FromSelector(idCount, classLikeCount, typeCount);
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
                        var max = MaxSpecificityFromSelectorList(arg);
                        idCount += max.Id;
                        classLikeCount += max.ClassLike;
                        typeCount += max.Type;
                    }
                }
                else if (!pseudoElement && pseudoName.Equals("is", StringComparison.OrdinalIgnoreCase) && i < simple.Length && simple[i] == '(')
                {
                    if (TryReadFunctionArgument(simple, ref i, out var arg))
                    {
                        var max = MaxSpecificityFromSelectorList(arg);
                        idCount += max.Id;
                        classLikeCount += max.ClassLike;
                        typeCount += max.Type;
                    }
                }
                else if (!pseudoElement && pseudoName.Equals("has", StringComparison.OrdinalIgnoreCase) && i < simple.Length && simple[i] == '(')
                {
                    if (TryReadFunctionArgument(simple, ref i, out var arg))
                    {
                        var max = MaxSpecificityFromSelectorList(arg, trimLeadingCombinator: true);
                        idCount += max.Id;
                        classLikeCount += max.ClassLike;
                        typeCount += max.Type;
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

    private static Specificity MaxSpecificityFromSelectorList(string selectorList, bool trimLeadingCombinator = false)
    {
        var max = Specificity.None;
        foreach (var part in SplitSelectorList(selectorList))
        {
            var trimmed = part.Trim();
            if (string.IsNullOrEmpty(trimmed)) continue;
            if (trimLeadingCombinator) trimmed = trimmed.TrimStart('>', '+', '~').Trim();
            if (string.IsNullOrEmpty(trimmed)) continue;
            var sp = CalculateSelectorSpecificity(trimmed);
            if (sp.CompareTo(max) > 0) max = sp;
        }
        return max;
    }

    private readonly struct Specificity : IComparable<Specificity>
    {
        public int InlineLevel { get; }
        public int Id { get; }
        public int ClassLike { get; }
        public int Type { get; }

        public static Specificity None => new(0, 0, 0, 0);
        public static Specificity InlineStyle => new(1, 0, 0, 0);

        public static Specificity FromSelector(int id, int classLike, int type) => new(0, id, classLike, type);

        public Specificity(int inline, int id, int classLike, int type)
        {
            InlineLevel = inline;
            Id = id;
            ClassLike = classLike;
            Type = type;
        }

        public int CompareTo(Specificity other)
        {
            if (InlineLevel != other.InlineLevel) return InlineLevel.CompareTo(other.InlineLevel);
            if (Id != other.Id) return Id.CompareTo(other.Id);
            if (ClassLike != other.ClassLike) return ClassLike.CompareTo(other.ClassLike);
            return Type.CompareTo(other.Type);
        }
    }

    private readonly struct StylePriority : IComparable<StylePriority>
    {
        public bool Important { get; }
        public Specificity Specificity { get; }
        public int SourceOrder { get; }

        public StylePriority(bool important, Specificity specificity, int sourceOrder)
        {
            Important = important;
            Specificity = specificity;
            SourceOrder = sourceOrder;
        }

        public int CompareTo(StylePriority other)
        {
            if (Important != other.Important) return Important ? 1 : -1;
            int sp = Specificity.CompareTo(other.Specificity);
            if (sp != 0) return sp;
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
                else if (pseudoName.Equals("root", StringComparison.OrdinalIgnoreCase) || pseudoName.Equals("scope", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsRootElement(node)) return false;
                }
                else if (pseudoName.Equals("lang", StringComparison.OrdinalIgnoreCase) && i < sel.Length && sel[i] == '(')
                {
                    if (!TryReadFunctionArgument(sel, ref i, out var arg)) return false;
                    if (!MatchesLang(node, arg)) return false;
                }
                else if (pseudoName.Equals("dir", StringComparison.OrdinalIgnoreCase) && i < sel.Length && sel[i] == '(')
                {
                    if (!TryReadFunctionArgument(sel, ref i, out var arg)) return false;
                    if (!MatchesDir(node, arg)) return false;
                }
                else if (pseudoName.Equals("checked", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsChecked(node)) return false;
                }
                else if (pseudoName.Equals("disabled", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsDisabled(node)) return false;
                }
                else if (pseudoName.Equals("enabled", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsEnabled(node)) return false;
                }
                else if (pseudoName.Equals("required", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsRequired(node)) return false;
                }
                else if (pseudoName.Equals("optional", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsOptional(node)) return false;
                }
                else if (pseudoName.Equals("read-only", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsReadOnly(node)) return false;
                }
                else if (pseudoName.Equals("read-write", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsReadWrite(node)) return false;
                }
                else if (pseudoName.Equals("link", StringComparison.OrdinalIgnoreCase) || pseudoName.Equals("any-link", StringComparison.OrdinalIgnoreCase))
                {
                    if (!IsAnyLink(node)) return false;
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
                if (MatchesRelativeSelector(node, neg)) return false;
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
        var tokens = TokenizeSelector(trimmed);
        if (tokens.Count == 0) return false;

        string startCombinator = " ";
        int index = 0;
        if (Combinators.Contains(tokens[0]))
        {
            startCombinator = tokens[0];
            index = 1;
        }

        if (index >= tokens.Count) return false;
        if (Combinators.Contains(tokens[^1])) return false;
        if (Combinators.Contains(tokens[index])) return false;

        var current = CollectRelated(scope, startCombinator)
            .Where(n => MatchesSimple(n, tokens[index]))
            .Distinct()
            .ToList();
        if (current.Count == 0) return false;

        index++;
        while (index < tokens.Count)
        {
            if (!Combinators.Contains(tokens[index])) return false;
            var combinator = tokens[index];
            index++;
            if (index >= tokens.Count || Combinators.Contains(tokens[index])) return false;
            var nextSimple = tokens[index];

            var next = new List<HtmlNode>();
            foreach (var n in current)
            {
                foreach (var related in CollectRelated(n, combinator))
                {
                    if (MatchesSimple(related, nextSimple))
                        next.Add(related);
                }
            }

            current = next.Distinct().ToList();
            if (current.Count == 0) return false;
            index++;
        }

        return current.Count > 0;
    }

    private static bool IsFirstChild(HtmlNode node)
    {
        var siblings = GetElementSiblings(node);
        return siblings.Count > 0 && ReferenceEquals(siblings[0], node);
    }

    private static bool IsRootElement(HtmlNode node)
    {
        return node.Parent != null && node.Parent.TagName == "#document";
    }

    private static bool MatchesLang(HtmlNode node, string arg)
    {
        var expected = arg.Trim().Trim('"', '\'');
        if (string.IsNullOrEmpty(expected)) return false;
        var actual = GetInheritedAttribute(node, "lang");
        if (string.IsNullOrEmpty(actual)) return false;
        return actual.Equals(expected, StringComparison.OrdinalIgnoreCase) ||
               actual.StartsWith(expected + "-", StringComparison.OrdinalIgnoreCase);
    }

    private static bool MatchesDir(HtmlNode node, string arg)
    {
        var expected = arg.Trim().Trim('"', '\'');
        if (string.IsNullOrEmpty(expected)) return false;
        var actual = GetInheritedAttribute(node, "dir");
        if (string.IsNullOrEmpty(actual)) return false;
        return actual.Equals(expected, StringComparison.OrdinalIgnoreCase);
    }

    private static string? GetInheritedAttribute(HtmlNode node, string name)
    {
        HtmlNode? current = node;
        while (current != null)
        {
            var value = current.GetAttribute(name);
            if (!string.IsNullOrWhiteSpace(value))
                return value;
            current = current.Parent;
        }
        return null;
    }

    private static bool IsChecked(HtmlNode node)
    {
        if (node.TagName.Equals("input", StringComparison.OrdinalIgnoreCase))
            return HasBooleanAttribute(node, "checked");
        if (node.TagName.Equals("option", StringComparison.OrdinalIgnoreCase))
            return HasBooleanAttribute(node, "selected") || HasBooleanAttribute(node, "checked");
        return false;
    }

    private static bool IsDisabled(HtmlNode node)
    {
        if (!IsFormControl(node)) return false;
        if (HasBooleanAttribute(node, "disabled")) return true;

        HtmlNode? current = node.Parent;
        while (current != null)
        {
            if (node.TagName.Equals("option", StringComparison.OrdinalIgnoreCase) &&
                current.TagName.Equals("optgroup", StringComparison.OrdinalIgnoreCase) &&
                HasBooleanAttribute(current, "disabled"))
                return true;

            if ((node.TagName.Equals("option", StringComparison.OrdinalIgnoreCase) || node.TagName.Equals("optgroup", StringComparison.OrdinalIgnoreCase)) &&
                current.TagName.Equals("select", StringComparison.OrdinalIgnoreCase) &&
                HasBooleanAttribute(current, "disabled"))
                return true;

            if (current.TagName.Equals("fieldset", StringComparison.OrdinalIgnoreCase) && HasBooleanAttribute(current, "disabled"))
            {
                if (!IsInsideFirstLegendOfFieldset(node, current))
                    return true;
            }
            current = current.Parent;
        }
        return false;
    }

    private static bool IsEnabled(HtmlNode node)
    {
        return IsFormControl(node) && !IsDisabled(node);
    }

    private static bool IsRequired(HtmlNode node)
    {
        return SupportsRequired(node) && HasBooleanAttribute(node, "required");
    }

    private static bool IsOptional(HtmlNode node)
    {
        return SupportsRequired(node) && !HasBooleanAttribute(node, "required");
    }

    private static bool IsReadOnly(HtmlNode node)
    {
        return !IsReadWrite(node);
    }

    private static bool IsReadWrite(HtmlNode node)
    {
        var contentEditable = node.GetAttribute("contenteditable");
        if (!string.IsNullOrWhiteSpace(contentEditable))
            return !contentEditable.Equals("false", StringComparison.OrdinalIgnoreCase);

        if (node.TagName.Equals("textarea", StringComparison.OrdinalIgnoreCase))
            return !HasBooleanAttribute(node, "readonly") && !IsDisabled(node);

        if (node.TagName.Equals("input", StringComparison.OrdinalIgnoreCase))
        {
            if (HasBooleanAttribute(node, "readonly") || IsDisabled(node)) return false;
            var type = (node.GetAttribute("type") ?? "text").Trim().ToLowerInvariant();
            return type != "hidden" &&
                   type != "button" &&
                   type != "submit" &&
                   type != "reset" &&
                   type != "checkbox" &&
                   type != "radio" &&
                   type != "file" &&
                   type != "image" &&
                   type != "color" &&
                   type != "range";
        }

        return false;
    }

    private static bool SupportsRequired(HtmlNode node)
    {
        if (node.TagName.Equals("select", StringComparison.OrdinalIgnoreCase) ||
            node.TagName.Equals("textarea", StringComparison.OrdinalIgnoreCase))
            return true;

        if (!node.TagName.Equals("input", StringComparison.OrdinalIgnoreCase))
            return false;

        var type = (node.GetAttribute("type") ?? "text").Trim().ToLowerInvariant();
        return type != "hidden" &&
               type != "button" &&
               type != "submit" &&
               type != "reset" &&
               type != "image" &&
               type != "color" &&
               type != "range";
    }

    private static bool IsAnyLink(HtmlNode node)
    {
        if (!HasAttribute(node, "href")) return false;
        return node.TagName.Equals("a", StringComparison.OrdinalIgnoreCase) ||
               node.TagName.Equals("area", StringComparison.OrdinalIgnoreCase) ||
               node.TagName.Equals("link", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsFormControl(HtmlNode node)
    {
        return node.TagName.Equals("input", StringComparison.OrdinalIgnoreCase) ||
               node.TagName.Equals("button", StringComparison.OrdinalIgnoreCase) ||
               node.TagName.Equals("select", StringComparison.OrdinalIgnoreCase) ||
               node.TagName.Equals("textarea", StringComparison.OrdinalIgnoreCase) ||
               node.TagName.Equals("option", StringComparison.OrdinalIgnoreCase) ||
               node.TagName.Equals("optgroup", StringComparison.OrdinalIgnoreCase) ||
               node.TagName.Equals("fieldset", StringComparison.OrdinalIgnoreCase);
    }

    private static bool HasBooleanAttribute(HtmlNode node, string name)
    {
        return node.Attributes.ContainsKey(name);
    }

    private static bool HasAttribute(HtmlNode node, string name)
    {
        return node.Attributes.ContainsKey(name);
    }

    private static bool IsInsideFirstLegendOfFieldset(HtmlNode node, HtmlNode fieldset)
    {
        var firstLegend = fieldset.Children.FirstOrDefault(c => c.TagName.Equals("legend", StringComparison.OrdinalIgnoreCase));
        if (firstLegend == null) return false;
        return IsDescendantOrSelf(node, firstLegend);
    }

    private static bool IsDescendantOrSelf(HtmlNode node, HtmlNode ancestor)
    {
        HtmlNode? current = node;
        while (current != null)
        {
            if (ReferenceEquals(current, ancestor)) return true;
            current = current.Parent;
        }
        return false;
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
        ParseNthExpressionAndSelector(expression, out var nthExpression, out var ofSelectorList);
        var siblings = GetNthCandidateSiblings(node, ofSelectorList);
        int index = siblings.IndexOf(node) + 1;
        if (index <= 0) return false;
        return MatchesNthExpression(index, nthExpression);
    }

    private static bool MatchesNthOfType(HtmlNode node, string expression)
    {
        ParseNthExpressionAndSelector(expression, out var nthExpression, out var ofSelectorList);
        var ofType = GetNthOfTypeCandidateSiblings(node, ofSelectorList);
        int index = ofType.IndexOf(node) + 1;
        if (index <= 0) return false;
        return MatchesNthExpression(index, nthExpression);
    }

    private static bool MatchesNthLastChild(HtmlNode node, string expression)
    {
        ParseNthExpressionAndSelector(expression, out var nthExpression, out var ofSelectorList);
        var siblings = GetNthCandidateSiblings(node, ofSelectorList);
        int indexFromEnd = siblings.Count - siblings.IndexOf(node);
        if (indexFromEnd <= 0) return false;
        return MatchesNthExpression(indexFromEnd, nthExpression);
    }

    private static List<HtmlNode> GetNthCandidateSiblings(HtmlNode node, string? ofSelectorList)
    {
        var siblings = GetElementSiblings(node);
        if (string.IsNullOrWhiteSpace(ofSelectorList))
            return siblings;

        return siblings.Where(s => MatchesAnySelectorInList(s, ofSelectorList)).ToList();
    }

    private static bool MatchesAnySelectorInList(HtmlNode node, string selectorList)
    {
        foreach (var part in SplitSelectorList(selectorList))
        {
            var selector = part.Trim();
            if (selector.Length == 0) continue;
            if (MatchesRelativeSelector(node, selector))
                return true;
        }
        return false;
    }

    private static void ParseNthExpressionAndSelector(string argument, out string nthExpression, out string? ofSelectorList)
    {
        nthExpression = argument.Trim();
        ofSelectorList = null;
        int bracketDepth = 0;
        int parenDepth = 0;
        char quote = '\0';
        for (int i = 0; i < argument.Length; i++)
        {
            char c = argument[i];
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

            if (char.IsWhiteSpace(c))
            {
                int j = i;
                while (j < argument.Length && char.IsWhiteSpace(argument[j])) j++;
                if (j + 1 < argument.Length &&
                    (argument[j] == 'o' || argument[j] == 'O') &&
                    (argument[j + 1] == 'f' || argument[j + 1] == 'F'))
                {
                    int k = j + 2;
                    if (k < argument.Length && char.IsWhiteSpace(argument[k]))
                    {
                        nthExpression = argument.Substring(0, i).Trim();
                        ofSelectorList = argument.Substring(k).Trim();
                        return;
                    }
                }
            }
        }
    }

    private static bool MatchesNthLastOfType(HtmlNode node, string expression)
    {
        ParseNthExpressionAndSelector(expression, out var nthExpression, out var ofSelectorList);
        var ofType = GetNthOfTypeCandidateSiblings(node, ofSelectorList);
        int indexFromEnd = ofType.Count - ofType.IndexOf(node);
        if (indexFromEnd <= 0) return false;
        return MatchesNthExpression(indexFromEnd, nthExpression);
    }

    private static List<HtmlNode> GetNthOfTypeCandidateSiblings(HtmlNode node, string? ofSelectorList)
    {
        var ofType = GetElementSiblings(node).Where(s => s.TagName.Equals(node.TagName, StringComparison.OrdinalIgnoreCase)).ToList();
        if (string.IsNullOrWhiteSpace(ofSelectorList))
            return ofType;

        return ofType.Where(s => MatchesAnySelectorInList(s, ofSelectorList)).ToList();
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

    private static IEnumerable<HtmlNode> CollectRelated(HtmlNode node, string combinator)
    {
        return combinator switch
        {
            ">" => node.Children.Where(c => c.TagName != "#text"),
            "+" => GetNextElementSibling(node) is HtmlNode n ? new[] { n } : Array.Empty<HtmlNode>(),
            "~" => GetFollowingElementSiblings(node),
            _ => EnumerateDescendants(node)
        };
    }

    private static IEnumerable<HtmlNode> EnumerateDescendants(HtmlNode node)
    {
        foreach (var child in node.Children)
        {
            if (child.TagName != "#text") yield return child;
            foreach (var inner in EnumerateDescendants(child)) yield return inner;
        }
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
