// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // SEQ field evaluator. Unlike PAGE/PAGEREF/TOC-page-numbers (which need a
    // layout/pagination engine officecli does not have), SEQ numbering is a pure
    // document-order count of same-identifier SEQ fields — no pagination involved,
    // so the value is computable from the document model alone. That lets us write
    // a correct cached <w:t> ourselves instead of emitting "1"/empty and deferring
    // everything to the application.
    //
    // Honesty rule (mirrors the xlsx evaluator's EvaluateForReport): only patch what
    // we can compute confidently — `\n` (next), `\c` (repeat current), `\r N` (reset),
    // and arabic / roman / alphabetic number formats, counted in BODY document order.
    // Anything we can't be sure of is LEFT UNTOUCHED for Word to recompute (combine
    // with `set /settings --prop updateFields=true`): `\s` heading-relative resets,
    // SEQ outside the main body (headers/footers/textboxes — different counting story),
    // and unsupported `\#` picture switches.

    private sealed class SeqInstr
    {
        public string Identifier = "";
        public long? Reset;          // \r N
        public bool RepeatCurrent;   // \c
        public string Format = "arabic"; // \* ROMAN | roman | ALPHABETIC | alphabetic | arabic
        public bool Defer;           // \s / unsupported \# / unparseable → leave to Word
    }

    /// <summary>
    /// Recompute and write cached values for SEQ fields in body document order.
    /// Returns the number of fields patched. Deferred fields (see Honesty rule)
    /// are left untouched so `evaluated` stays false and Word fills them on open.
    /// </summary>
    internal int RecalcSeqFields()
    {
        var body = _doc.MainDocumentPart?.Document?.Body;
        if (body == null) return 0;

        var fields = new List<FieldInfo>();
        CollectFieldsFrom(body.Descendants<Run>(), fields, body); // body, document order

        var counters = new Dictionary<string, long>(StringComparer.Ordinal);
        int patched = 0;
        foreach (var f in fields)
        {
            var instr = f.InstrCode.Text?.Trim() ?? "";
            if (instr.Split(' ', 2)[0].ToUpperInvariant() != "SEQ") continue;
            var p = ParseSeqInstruction(instr);
            if (p == null || p.Defer) continue;

            long cur = counters.GetValueOrDefault(p.Identifier, 0);
            long val;
            if (p.Reset.HasValue) val = p.Reset.Value;
            else if (p.RepeatCurrent) { if (cur == 0) continue; val = cur; }
            else val = cur + 1;
            counters[p.Identifier] = val;

            var text = FormatSeqValue(val, p.Format);
            if (text == null) continue; // format we don't render → leave to Word
            SetFieldResultText(f, text);
            patched++;
        }
        if (patched > 0) SaveDoc();
        return patched;
    }

    private static SeqInstr? ParseSeqInstruction(string instr)
    {
        var toks = TokenizeFieldInstruction(instr);
        if (toks.Count < 2) return null; // need SEQ + identifier
        var p = new SeqInstr { Identifier = toks[1] };
        for (int i = 2; i < toks.Count; i++)
        {
            var t = toks[i];
            if (!t.StartsWith('\\')) continue;
            switch (t.ToLowerInvariant())
            {
                case "\\n": break;                 // next (default)
                case "\\c": p.RepeatCurrent = true; break;
                case "\\h": break;                 // hidden result still counts
                case "\\r":
                    if (i + 1 < toks.Count && long.TryParse(toks[++i], out var rn)) p.Reset = rn;
                    else p.Defer = true;
                    break;
                case "\\s":                        // heading-relative reset — needs heading walk
                    p.Defer = true;
                    if (i + 1 < toks.Count && !toks[i + 1].StartsWith('\\')) i++;
                    break;
                case "\\*":
                    if (i + 1 < toks.Count && !toks[i + 1].StartsWith('\\'))
                    {
                        var fmt = toks[++i];
                        p.Format = fmt switch
                        {
                            "ROMAN" => "ROMAN",
                            "roman" => "roman",
                            "ALPHABETIC" => "ALPHABETIC",
                            "alphabetic" => "alphabetic",
                            "Arabic" or "arabic" => "arabic",
                            _ => "defer", // CardText / Ordinal / Hex / … not rendered yet
                        };
                        if (p.Format == "defer") p.Defer = true;
                    }
                    break;
                case "\\#":                        // numeric picture — only honour if absent
                    p.Defer = true;                // \# "0.00" etc. → let Word render
                    if (i + 1 < toks.Count && !toks[i + 1].StartsWith('\\')) i++;
                    break;
                default: break;                    // unknown switch — ignore, still count
            }
        }
        return p;
    }

    // Split a field instruction into tokens, treating a "quoted string" as one
    // token and stripping the quotes (matches Word's field-arg tokenizing).
    private static List<string> TokenizeFieldInstruction(string instr)
    {
        var toks = new List<string>();
        int i = 0;
        while (i < instr.Length)
        {
            while (i < instr.Length && char.IsWhiteSpace(instr[i])) i++;
            if (i >= instr.Length) break;
            if (instr[i] == '"')
            {
                int start = ++i;
                while (i < instr.Length && instr[i] != '"') i++;
                toks.Add(instr.Substring(start, i - start));
                if (i < instr.Length) i++; // skip closing quote
            }
            else
            {
                int start = i;
                while (i < instr.Length && !char.IsWhiteSpace(instr[i])) i++;
                toks.Add(instr.Substring(start, i - start));
            }
        }
        return toks;
    }

    private static string? FormatSeqValue(long n, string format) => format switch
    {
        "arabic" => n.ToString(System.Globalization.CultureInfo.InvariantCulture),
        "ROMAN" => n is > 0 and < 4000 ? ToRoman((int)n) : null,
        "roman" => n is > 0 and < 4000 ? ToRoman((int)n).ToLowerInvariant() : null,
        "ALPHABETIC" => n > 0 ? ToAlpha((int)n, upper: true) : null,
        "alphabetic" => n > 0 ? ToAlpha((int)n, upper: false) : null,
        _ => null,
    };

    private static string ToRoman(int n)
    {
        int[] vals = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        string[] sym = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < vals.Length && n > 0; i++)
            while (n >= vals[i]) { sb.Append(sym[i]); n -= vals[i]; }
        return sb.ToString();
    }

    // Word `\* alphabetic`: 1→A, 26→Z, 27→AA, 28→BB (each letter repeated), matching
    // Word's spreadsheet-unlike scheme where 27 = "AA", 53 = "AAA".
    private static string ToAlpha(int n, bool upper)
    {
        int count = (n - 1) / 26 + 1;
        char letter = (char)('A' + (n - 1) % 26);
        if (!upper) letter = char.ToLowerInvariant(letter);
        return new string(letter, count);
    }

    // Write <paramref name="text"/> as the field's cached result, creating the
    // separate / result runs if the field lacks them.
    private static void SetFieldResultText(FieldInfo f, string text)
    {
        Run target;
        if (f.ResultRuns.Count > 0)
        {
            target = f.ResultRuns[0];
            foreach (var c in target.ChildElements.Where(c => c is not RunProperties).ToList())
                c.Remove();
        }
        else
        {
            target = new Run();
            if (f.SeparateRun != null)
                f.SeparateRun.InsertAfterSelf(target);
            else
            {
                // No separator run: synthesize begin…instr…[separate][result]…end.
                var instrRun = f.InstrCode.Parent as Run;
                var sep = new Run(new FieldChar { FieldCharType = FieldCharValues.Separate });
                if (instrRun != null) { instrRun.InsertAfterSelf(sep); sep.InsertAfterSelf(target); }
                else f.EndRun.InsertBeforeSelf(target);
            }
        }
        target.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }
}
