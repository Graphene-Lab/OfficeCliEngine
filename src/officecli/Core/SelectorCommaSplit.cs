// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Top-level comma handling for CSS-style selector lists (union). A comma is
/// "top level" only when it sits outside every <c>[]</c> / <c>()</c> group and
/// outside quotes, so a value like <c>row[name="A,B"]</c> or a function arg is
/// never split. Shared by the PowerPoint and Excel query paths so comma-union
/// behaves identically across handlers (CONSISTENCY(comma-union)).
/// </summary>
internal static class SelectorCommaSplit
{
    public static bool ContainsTopLevelComma(string selector)
        => ContainsTopLevelChar(selector, ',');

    /// <summary>
    /// True when <paramref name="target"/> appears outside every <c>[]</c> /
    /// <c>()</c> group and outside quotes (a backslash escapes the next char
    /// inside a quoted span). Used for ',' (union split) and '/' (detecting a
    /// slash path that lost its leading slash) — same scan, one impl.
    /// </summary>
    public static bool ContainsTopLevelChar(string selector, char target)
        => TopLevelIndexOf(selector, target) >= 0;

    /// <summary>
    /// Index of the first top-level occurrence of <paramref name="target"/>,
    /// or -1. Same scan as <see cref="ContainsTopLevelChar"/>.
    /// </summary>
    public static int TopLevelIndexOf(string selector, char target)
    {
        int depthBracket = 0, depthParen = 0;
        char? quote = null;
        for (int i = 0; i < selector.Length; i++)
        {
            var c = selector[i];
            if (quote.HasValue)
            {
                if (c == '\\' && i + 1 < selector.Length) { i++; continue; }
                if (c == quote.Value) quote = null;
                continue;
            }
            if (c == '"' || c == '\'') { quote = c; continue; }
            if (c == '[') depthBracket++;
            else if (c == ']') depthBracket = System.Math.Max(0, depthBracket - 1);
            else if (c == '(') depthParen++;
            else if (c == ')') depthParen = System.Math.Max(0, depthParen - 1);
            else if (c == target && depthBracket == 0 && depthParen == 0) return i;
        }
        return -1;
    }

    public static List<string> SplitTopLevelCommas(string selector)
    {
        var parts = new List<string>();
        int depthBracket = 0, depthParen = 0;
        char? quote = null;
        int start = 0;
        for (int i = 0; i < selector.Length; i++)
        {
            var c = selector[i];
            if (quote.HasValue)
            {
                if (c == '\\' && i + 1 < selector.Length) { i++; continue; }
                if (c == quote.Value) quote = null;
                continue;
            }
            if (c == '"' || c == '\'') { quote = c; continue; }
            if (c == '[') depthBracket++;
            else if (c == ']') depthBracket = System.Math.Max(0, depthBracket - 1);
            else if (c == '(') depthParen++;
            else if (c == ')') depthParen = System.Math.Max(0, depthParen - 1);
            else if (c == ',' && depthBracket == 0 && depthParen == 0)
            {
                parts.Add(selector.Substring(start, i - start));
                start = i + 1;
            }
        }
        parts.Add(selector.Substring(start));
        return parts;
    }
}
