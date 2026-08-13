// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Selector ====================

    private record CellSelector(string? Sheet, string? Column, string? ValueEquals, string? ValueNotEquals,
        string? ValueContains, bool? HasFormula, bool? IsEmpty, string? TypeEquals, string? TypeNotEquals,
        Dictionary<string, string>? FormatEquals = null, Dictionary<string, string>? FormatNotEquals = null);

    private CellSelector ParseCellSelector(string selector)
    {
        string? sheet = null;
        string? column = null;
        string? valueEquals = null;
        string? valueNotEquals = null;
        string? valueContains = null;
        bool? hasFormula = null;
        bool? isEmpty = null;
        string? typeEquals = null;
        string? typeNotEquals = null;

        // Normalize path-style selectors: "/Sheet1/cell[...]" → "Sheet1!cell[...]"
        if (selector.StartsWith('/'))
        {
            var slashIdx = selector.IndexOf('/', 1);
            if (slashIdx > 0)
            {
                sheet = selector[1..slashIdx];
                selector = selector[(slashIdx + 1)..];
            }
            else
            {
                // Just "/cell" — strip leading slash
                selector = selector[1..];
            }
        }

        // Check for sheet prefix: Sheet1!cell[...]
        // Only a TOP-LEVEL '!' (outside brackets/quotes, not part of '!=')
        // separates the sheet — a '!' inside a predicate value
        // (row[F="x!"], row[Msg~=hello!]) is literal text, and the old
        // ^(.+?)!(?!=) regex swallowed everything up to it as a "sheet".
        var bangIdx = Core.SelectorCommaSplit.TopLevelIndexOf(selector, '!');
        if (bangIdx > 0 && (bangIdx + 1 >= selector.Length || selector[bangIdx + 1] != '='))
        {
            // Excel-quoted names ('My Data (2024)'!row[...]) arrive quoted —
            // the scanner already treats the quoted span as opaque, so the
            // top-level '!' is the real separator; strip the quotes here.
            sheet = UnquoteSheetName(selector[..bangIdx]);
            selector = selector[(bangIdx + 1)..];
        }

        // Parse element and attributes: cell[attr=value]
        var match = Regex.Match(selector, @"^(\w+)?(.*)$");
        var element = match.Groups[1].Value;

        // Column filter: e.g., "B" or "AB" — but NOT known element types like "row"
        if (element.Length <= 3 && Regex.IsMatch(element, @"^[A-Z]+$", RegexOptions.IgnoreCase)
            && element.ToLowerInvariant() is not ("row" or "cell" or "col"))
        {
            column = element.ToUpperInvariant();
        }

        // Parse attributes (\\?! handles zsh escaping \! as !)
        Dictionary<string, string>? formatEquals = null;
        Dictionary<string, string>? formatNotEquals = null;
        foreach (Match attrMatch in Regex.Matches(selector, @"\[([\w.]+)(\\?!?=)([^\]]*)\]"))
        {
            var key = attrMatch.Groups[1].Value.ToLowerInvariant();
            var op = attrMatch.Groups[2].Value.Replace("\\", "");
            var val = attrMatch.Groups[3].Value.Trim('\'', '"');

            switch (key)
            {
                case "value" when op == "=": valueEquals = val; break;
                case "value" when op == "!=": valueNotEquals = val; break;
                case "type" when op == "=": typeEquals = val; break;
                case "type" when op == "!=": typeNotEquals = val; break;
                case "formula": hasFormula = val.ToLowerInvariant() != "false"; break;
                case "empty": isEmpty = val.ToLowerInvariant() != "false"; break;
                default:
                    if (op == "=")
                    {
                        formatEquals ??= new Dictionary<string, string>();
                        formatEquals[attrMatch.Groups[1].Value] = val;
                    }
                    else if (op == "!=")
                    {
                        formatNotEquals ??= new Dictionary<string, string>();
                        formatNotEquals[attrMatch.Groups[1].Value] = val;
                    }
                    break;
            }
        }

        // :contains() pseudo-selector
        var containsMatch = Regex.Match(selector, @":contains\(['""]?(.+?)['""]?\)");
        if (containsMatch.Success) valueContains = containsMatch.Groups[1].Value;

        // Shorthand: "cell:text" → treat as :contains(text). Exclude the
        // recognised pseudo-classes (incl. `not`) so `:not(:has(formula))` is
        // not mistaken for a literal substring filter "not(:has(formula))".
        if (valueContains == null)
        {
            var shorthandMatch = Regex.Match(selector, @"^(?:\w+)?:(?!contains|empty|has|not)(.+)$");
            if (shorthandMatch.Success) valueContains = shorthandMatch.Groups[1].Value;
        }

        // :empty pseudo-selector (and its negation). Check the negated form
        // first — `:not(:empty)` also contains the `:empty` substring.
        if (selector.Contains(":not(:empty)")) isEmpty = false;
        else if (selector.Contains(":empty")) isEmpty = true;

        // :has(formula) pseudo-selector (and its negation). `:not(:has(formula))`
        // contains `:has(formula)`, so the negated form must be checked first —
        // otherwise it inverts to "has a formula" and returns the wrong set.
        if (selector.Contains(":not(:has(formula))")) hasFormula = false;
        else if (selector.Contains(":has(formula)")) hasFormula = true;

        return new CellSelector(sheet, column, valueEquals, valueNotEquals, valueContains, hasFormula, isEmpty, typeEquals, typeNotEquals, formatEquals, formatNotEquals);
    }

    private bool MatchesCellSelector(Cell cell, string sheetName, CellSelector selector)
    {
        // Column filter
        if (selector.Column != null)
        {
            var cellRef = cell.CellReference?.Value ?? "";
            var (colName, _) = ParseCellReference(cellRef);
            if (!colName.Equals(selector.Column, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        var value = GetCellDisplayValue(cell);
        // Stored value for a formatted cell (0.5, not "50%"; date serial, not the
        // formatted date). Equality matches EITHER form so `value=50%` (display)
        // and `value=0.5` (stored) both hit the same percentage cell.
        var rawValue = GetCellRawComparisonValue(cell);
        bool ValueMatches(string target) =>
            value.Equals(target, StringComparison.OrdinalIgnoreCase)
            || rawValue.Equals(target, StringComparison.OrdinalIgnoreCase);

        // Value filters
        if (selector.ValueEquals != null && !ValueMatches(selector.ValueEquals))
            return false;
        if (selector.ValueNotEquals != null && ValueMatches(selector.ValueNotEquals))
            return false;
        if (selector.ValueContains != null && !value.Contains(selector.ValueContains, StringComparison.OrdinalIgnoreCase))
            return false;

        // Formula filter
        if (selector.HasFormula == true && cell.CellFormula == null)
            return false;
        if (selector.HasFormula == false && cell.CellFormula != null)
            return false;

        // Empty filter
        if (selector.IsEmpty == true && !string.IsNullOrEmpty(value))
            return false;
        if (selector.IsEmpty == false && string.IsNullOrEmpty(value))
            return false;

        // Type filter (use friendly names matching CellToNode output)
        if (selector.TypeEquals != null || selector.TypeNotEquals != null)
        {
            var type = GetCellTypeName(cell);
            if (selector.TypeEquals != null && !type.Equals(selector.TypeEquals, StringComparison.OrdinalIgnoreCase))
                return false;
            if (selector.TypeNotEquals != null && type.Equals(selector.TypeNotEquals, StringComparison.OrdinalIgnoreCase))
                return false;
        }

        return true;
    }

    private static string GetCellTypeName(Cell cell)
    {
        if (cell.DataType?.HasValue != true) return "Number";
        var dt = cell.DataType.Value;
        if (dt == CellValues.String) return "String";
        if (dt == CellValues.SharedString) return "SharedString";
        if (dt == CellValues.Boolean) return "Boolean";
        if (dt == CellValues.Error) return "Error";
        if (dt == CellValues.InlineString) return "InlineString";
        if (dt == CellValues.Date) return "Date";
        return "Number";
    }

    // CONSISTENCY(cell-selector-alias): short attribute names in cell selectors
    // map to their canonical DocumentNode.Format keys. Users write
    // `cell[bold=true]` but Get stores `font.bold`.
    private static readonly Dictionary<string, string> _cellSelectorAliases =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["bold"] = "font.bold",
            ["italic"] = "font.italic",
            // `strike`/`underline` are themselves the canonical cell Format keys
            // (Get emits them unprefixed), so they map to identity — no remap.
            ["underline"] = "underline",
            ["strike"] = "strike",
            ["font"] = "font.name",
            ["size"] = "font.size",
            ["color"] = "font.color",
        };

    private static string ResolveCellFormatKey(string key)
        => _cellSelectorAliases.TryGetValue(key, out var canonical) ? canonical : key;

    // CONSISTENCY(cell-selector-alias): exposed so the CLI query post-filter
    // (AttributeFilter.ApplyWithWarnings) can normalize user-written keys like
    // "bold" -> "font.bold" before matching against DocumentNode.Format. Without
    // this, handler-level MatchesCellSelector would accept cell[bold=true] and
    // return hits, then the CLI post-filter would drop them all because Format
    // only has "font.bold".
    public static string ResolveCellAttributeAlias(string key)
        => _cellSelectorAliases.TryGetValue(key, out var canonical) ? canonical : key;

    // CONSISTENCY(cell-selector-alias): true when a query selector targets cells,
    // accounting for an optional sheet prefix — bare "cell[...]", Excel-native
    // "Sheet1!cell[...]", or path-style "/Sheet1/cell[...]". The CLI and resident
    // query post-filters use this to decide whether to apply cell attribute-alias
    // normalization (bold -> font.bold, ...). Without stripping the prefix, a
    // sheet-scoped cell selector skipped normalization and every format-attribute
    // filter silently dropped all matches.
    public static bool SelectorTargetsCells(string? selector)
    {
        var element = Regex.Replace((selector ?? "").TrimStart(), @"^(?:[^/!\[]+!|/[^/]+/)", "");
        return element.StartsWith("cell", StringComparison.OrdinalIgnoreCase);
    }

    private static bool MatchesFormatAttributes(DocumentNode node, CellSelector selector)
    {
        if (selector.FormatEquals != null)
        {
            foreach (var (rawKey, expected) in selector.FormatEquals)
            {
                var key = ResolveCellFormatKey(rawKey);
                var matchedKey = node.Format.Keys.FirstOrDefault(k => string.Equals(k, key, StringComparison.OrdinalIgnoreCase));
                if (matchedKey == null) return false;
                var actual = node.Format[matchedKey]?.ToString() ?? "";
                if (!ColorNormalizedEquals(actual, expected))
                    return false;
            }
        }
        if (selector.FormatNotEquals != null)
        {
            foreach (var (rawKey, expected) in selector.FormatNotEquals)
            {
                var key = ResolveCellFormatKey(rawKey);
                var matchedKey = node.Format.Keys.FirstOrDefault(k => string.Equals(k, key, StringComparison.OrdinalIgnoreCase));
                var actual = matchedKey != null ? (node.Format[matchedKey]?.ToString() ?? "") : "";
                if (ColorNormalizedEquals(actual, expected))
                    return false;
            }
        }
        return true;
    }

    /// <summary>
    /// Compare two strings with color-aware normalization: "#FF0000" matches "FF0000".
    /// </summary>
    private static bool ColorNormalizedEquals(string a, string b)
    {
        if (string.Equals(a, b, StringComparison.OrdinalIgnoreCase)) return true;
        return string.Equals(a.TrimStart('#'), b.TrimStart('#'), StringComparison.OrdinalIgnoreCase);
    }

    // ==================== Cell Reference Utils ====================

    private static (string Column, int Row) ParseCellReference(string cellRef)
    {
        var match = Regex.Match(cellRef, @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
        if (!match.Success)
            throw new ArgumentException($"Invalid cell reference: '{cellRef}'. Expected format like 'A1', 'B2', 'XFD1048576'.");
        var col = match.Groups[1].Value.ToUpperInvariant();
        // Use long to avoid OverflowException when malformed files carry row numbers
        // outside int range (e.g. uint.MaxValue). Surface a semantic ArgumentException
        // (the same exception type used for other invalid refs below) instead.
        if (!long.TryParse(match.Groups[2].Value, out var rowLong) || rowLong < 1 || rowLong > 1048576)
            throw new ArgumentException(
                $"Row {match.Groups[2].Value} in cell reference '{cellRef}' is out of valid range. Valid range: 1-1048576.");
        var row = (int)rowLong;
        var colIdx = ColumnNameToIndex(col);
        if (colIdx < 1 || colIdx > 16384)
            throw new ArgumentException($"Column '{col}' in cell reference '{cellRef}' is out of range. Valid range: A-XFD (1-16384).");
        return (col, row);
    }

    private static int ColumnNameToIndex(string col)
    {
        int result = 0;
        foreach (var c in col.ToUpperInvariant())
        {
            result = result * 26 + (c - 'A' + 1);
        }
        return result;
    }

    private static string IndexToColumnName(int index)
    {
        var result = "";
        while (index > 0)
        {
            index--;
            result = (char)('A' + index % 26) + result;
            index /= 26;
        }
        return result;
    }

    private static DocumentFormat.OpenXml.Packaging.ChartPart GetChartPart(WorksheetPart worksheetPart, int index)
    {
        var drawingsPart = worksheetPart.DrawingsPart
            ?? throw new ArgumentException("Sheet has no drawings/charts");
        var chartParts = drawingsPart.ChartParts.ToList();
        if (index < 1 || index > chartParts.Count)
            throw new ArgumentException($"Chart index {index} out of range (1..{chartParts.Count})");
        return chartParts[index - 1];
    }

    private DocumentFormat.OpenXml.Packaging.ChartPart GetGlobalChartPart(int index)
    {
        var allCharts = new List<DocumentFormat.OpenXml.Packaging.ChartPart>();
        foreach (var (_, worksheetPart) in GetWorksheets())
        {
            if (worksheetPart.DrawingsPart != null)
                allCharts.AddRange(worksheetPart.DrawingsPart.ChartParts);
        }
        if (allCharts.Count == 0)
            throw new ArgumentException("No charts found in workbook");
        if (index < 1 || index > allCharts.Count)
            throw new ArgumentException($"Chart index {index} out of range (1..{allCharts.Count})");
        return allCharts[index - 1];
    }

    /// <summary>Charts addressable by a raw <c>chart[N]</c> path: the anchored
    /// charts (drawing order, both legacy and cx — matching query/get), followed
    /// by any chart parts NOT yet referenced by a graphicFrame. The trailing
    /// unanchored parts restore the <c>add-part → raw-set chart[N]</c> workflow,
    /// where the chart XML is authored before its drawing anchor exists.</summary>
    private static List<ExcelChartInfo> ChartsForRaw(DrawingsPart drawingsPart)
    {
        var list = GetExcelCharts(drawingsPart);
        var seen = new HashSet<OpenXmlPart>();
        foreach (var c in list)
            seen.Add(c.IsExtended ? (OpenXmlPart)c.ExtendedPart! : c.StandardPart!);
        foreach (var cp in drawingsPart.ChartParts)
            if (seen.Add(cp)) list.Add(new ExcelChartInfo { StandardPart = cp });
        foreach (var ep in drawingsPart.ExtendedChartParts)
            if (seen.Add(ep)) list.Add(new ExcelChartInfo { ExtendedPart = ep });
        return list;
    }

    /// <summary>Raw-XML resolver for a sheet-scoped chart path (/Sheet/chart[N]).
    /// Resolves anchored (drawing-order) and unanchored (add-part) charts, both
    /// legacy and extended (cx).</summary>
    private static string GetChartSpaceOuterXml(DrawingsPart? drawingsPart, int index)
    {
        if (drawingsPart == null)
            throw new ArgumentException("Sheet has no drawings/charts");
        return ChartSpaceOuterXmlAt(ChartsForRaw(drawingsPart), index);
    }

    /// <summary>Return the ChartSpace OuterXml of the 1-based chart in <paramref name="charts"/>,
    /// picking the legacy or extended root as appropriate.</summary>
    private static string ChartSpaceOuterXmlAt(List<ExcelChartInfo> charts, int index)
    {
        if (index < 1 || index > charts.Count)
            throw new ArgumentException($"Chart index {index} out of range (1..{charts.Count})");
        var info = charts[index - 1];
        return info.IsExtended
            ? info.ExtendedPart!.ChartSpace!.OuterXml
            : info.StandardPart!.ChartSpace!.OuterXml;
    }

    /// <summary>Live ChartSpace root for raw-set writes on a sheet-scoped chart
    /// path. Resolves both legacy and extended (cx) charts (see GetExcelCharts).</summary>
    private static DocumentFormat.OpenXml.OpenXmlPartRootElement GetChartSpaceElement(DrawingsPart? drawingsPart, int index)
    {
        if (drawingsPart == null)
            throw new ArgumentException("Sheet has no drawings/charts");
        return ChartSpaceElementAt(ChartsForRaw(drawingsPart), index);
    }

    /// <summary>Live ChartSpace root of the 1-based chart in <paramref name="charts"/>,
    /// legacy or extended.</summary>
    private static DocumentFormat.OpenXml.OpenXmlPartRootElement ChartSpaceElementAt(List<ExcelChartInfo> charts, int index)
    {
        if (index < 1 || index > charts.Count)
            throw new ArgumentException($"Chart index {index} out of range (1..{charts.Count})");
        var info = charts[index - 1];
        return info.IsExtended
            ? info.ExtendedPart!.ChartSpace!
            : info.StandardPart!.ChartSpace!;
    }
}
