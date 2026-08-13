// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// The project's ONE edit-distance implementation, backing every
/// "did you mean" suggester (property names, schema help keys, selector
/// attribute keys, table column names). Optimal string alignment
/// (restricted Damerau-Levenshtein): insert/delete/substitute cost 1, and a
/// swap of two ADJACENT characters also costs 1 — so the most common
/// real-world typo class (`blod`→`bold`, `Salray`→`Salary`) scores 1 instead
/// of Levenshtein's 2 and stays inside tight suggestion thresholds.
/// Case-sensitive; callers lowercase both sides first (every suggester wants
/// case-insensitive ranking).
/// </summary>
internal static class EditDistance
{
    public static int Damerau(string s, string t)
    {
        if (s.Length == 0) return t.Length;
        if (t.Length == 0) return s.Length;
        var d = new int[s.Length + 1, t.Length + 1];
        for (int i = 0; i <= s.Length; i++) d[i, 0] = i;
        for (int j = 0; j <= t.Length; j++) d[0, j] = j;
        for (int i = 1; i <= s.Length; i++)
        {
            for (int j = 1; j <= t.Length; j++)
            {
                int cost = s[i - 1] == t[j - 1] ? 0 : 1;
                d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
                if (i > 1 && j > 1 && s[i - 1] == t[j - 2] && s[i - 2] == t[j - 1])
                    d[i, j] = Math.Min(d[i, j], d[i - 2, j - 2] + 1);
            }
        }
        return d[s.Length, t.Length];
    }
}
