// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    // ==================== Column-predicate row matching ====================
    //
    // `row[Salary>5000]` matches the DATA rows of a table by a column's cell
    // value, addressing the column by its header NAME (or column letter). This
    // is the human/agent-natural "where" form for a normal table — you think in
    // column names, not B/E. It reuses the shared AttributeFilter operator
    // engine for comparison and the ListObject column metadata (names + range +
    // header/totals flags) already surfaced by TableToNode, so no header
    // sniffing or new comparison logic is introduced. P1 scope: real
    // ListObjects only; auto-binds to the single table that owns the
    // referenced column(s). Returns row nodes (`/Sheet/row[N]`) so the result
    // feeds Set/Remove for "match rows then operate".

    // Bracket keys that filter row PROPERTIES (height/hidden/...). Any other key
    // in a row selector is treated as a table COLUMN reference.
    private static readonly HashSet<string> RowAttributeKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "height", "hidden", "outlineLevel", "collapsed", "customHeight",
    };

    // A detected table's column names. Prefer the structured list carried on
    // InternalFormat (comma-safe); fall back to splitting the display string
    // only for nodes built before that carrier existed. Splitting Format
    // ["columns"] directly is WRONG when a header contains a comma
    // ("Amount, USD") — it invents a phantom column and shifts every header to
    // its right, silently mis-resolving `row[HeaderName op val]`.
    private static List<string> DetectedTableColumns(DocumentNode det)
    {
        if (det.InternalFormat.TryGetValue("columnList", out var lv) && lv is List<string> list)
            return list;
        return (det.Format.TryGetValue("columns", out var cv) ? cv?.ToString() ?? "" : "")
            .Split(',').Select(s => s.Trim()).ToList();
    }

    // Strip a `col.` / `column.` namespace prefix from a predicate key. The
    // prefix forces COLUMN interpretation at parse time; the resolver works on
    // the bare name. Case-insensitive. Returns the key unchanged when absent.
    private static string StripColPrefix(string key)
    {
        var m = Regex.Match(key, @"^col(?:umn)?\.(.+)$", RegexOptions.IgnoreCase);
        return m.Success ? m.Groups[1].Value : key;
    }


    // Match table data rows whose cells satisfy every column predicate. Auto-
    // binds to the single ListObject that owns all referenced columns; throws on
    // no-match or cross-table ambiguity rather than silently picking one.
    private List<DocumentNode> QueryRowsByColumnPredicate(string? sheetFilter, AttributeFilter.FilterExpr expr)
    {
        // Distinct leaf predicates drive column resolution and probe building; the
        // expression tree (which may be `or` / parens, e.g. row[金额>5 or 金额<1])
        // drives the actual row match via MatchesExpr.
        var colConds = AttributeFilter.LeafConditions(expr).ToList();
        // `*/row[...]` reads as a literal sheet named "*" and would report the
        // baffling "no table on sheet '*'". Wildcard sheet scope is not a
        // thing — the UNSCOPED form already searches every sheet.
        if (sheetFilter == "*")
            throw new Core.CliException(
                "Wildcard sheet scope ('*/row[...]') is not supported — omit the sheet prefix entirely; " +
                "an unscoped row[...] searches every sheet.")
            { Code = "invalid_selector", Suggestion = "row[...] (no sheet prefix) searches all sheets" };
        // A table that owns every referenced column. ListObjects are
        // authoritative (column names are stored metadata); detected tables are
        // a header-sniff heuristic (stable=false), so they are only consulted
        // when no ListObject matches — a real table never loses to a guess.
        var listObjCands = new List<(string sheetName, WorksheetPart part, string label, string source,
            int dataR1, int dataR2, Dictionary<string, int> colAbsIndex)>();
        var detectedCands = new List<(string sheetName, WorksheetPart part, string label, string source,
            int dataR1, int dataR2, Dictionary<string, int> colAbsIndex)>();
        // Every table in scope with its header names — used only to build a
        // helpful "no such column" error (list all when few, suggest similar
        // when many). Collected regardless of whether a table resolves the
        // predicate, so a typo'd column still surfaces the real names.
        var scopeTables = new List<(string label, List<string> cols)>();

        foreach (var (sheetName, worksheetPart) in GetWorksheets())
        {
            if (sheetFilter != null && !sheetName.Equals(sheetFilter, StringComparison.OrdinalIgnoreCase))
                continue;

            // ListObjects (authoritative column names).
            var realRanges = new List<(int c1, int r1, int c2, int r2)>();
            foreach (var tdp in worksheetPart.TableDefinitionParts)
            {
                var tbl = tdp.Table;
                if (tbl?.Reference?.Value == null) continue;
                if (!TryParseRange(tbl.Reference.Value, out var rng)) continue;
                realRanges.Add(rng);

                var colNames = tbl.GetFirstChild<TableColumns>()?.Elements<TableColumn>()
                    .Select(c => c.Name?.Value ?? "").ToList() ?? new List<string>();
                scopeTables.Add((tbl.Name?.Value ?? "table", colNames));
                var resolved = ResolveColumns(colNames, rng.c1, rng.c2, colConds);
                if (resolved == null) continue;

                bool headerRow = (tbl.HeaderRowCount?.Value ?? 1) != 0;
                bool totalRow = (tbl.TotalsRowCount?.Value ?? 0) > 0 || (tbl.TotalsRowShown?.Value ?? false);
                listObjCands.Add((sheetName, worksheetPart, tbl.Name?.Value ?? "table", "table",
                    rng.r1 + (headerRow ? 1 : 0), rng.r2 - (totalRow ? 1 : 0), resolved));
            }

            // Detected tables (header-sniff). Header is the first row of ref, no
            // totals row; data rows are ref minus the header row.
            foreach (var det in DetectTables(sheetName, worksheetPart, realRanges))
            {
                var colNames = DetectedTableColumns(det);
                var refStr = det.Format.TryGetValue("ref", out var rv) ? rv?.ToString() : null;
                scopeTables.Add((refStr ?? "detected", colNames));
                if (!TryParseRange(refStr, out var frng)) continue;
                var resolved = ResolveColumns(colNames, frng.c1, frng.c2, colConds);
                if (resolved == null) continue;
                detectedCands.Add((sheetName, worksheetPart, refStr!, "detected",
                    frng.r1 + 1, frng.r2, resolved));
            }
        }

        var candidates = listObjCands.Count > 0 ? listObjCands : detectedCands;

        if (candidates.Count == 0)
        {
            var cols = string.Join(", ", colConds.Select(c => $"'{StripColPrefix(c.Key)}'"));
            var scope = sheetFilter == null ? "any sheet" : $"sheet '{sheetFilter}'";
            // If NO table structure exists in scope at all, but the column NAME
            // literally exists as a cell, the honest cause is "not a recognizable
            // table" (a blank-header gap column, or a header-only block with no
            // data rows) — point at explicit table creation. Guard on an empty
            // scopeTables: when a real table WAS detected but simply doesn't own
            // one of a compound predicate's columns (`row[Region=X and Bonus>1]`
            // where Bonus is absent), fall through to BuildNoColumnException so
            // the error names the missing column and lists the available ones,
            // instead of wrongly blaming a valid column's surrounding cells.
            if (scopeTables.Count == 0 && FindHeaderLikeCell(sheetFilter, colConds) is {} h)
                throw new Core.CliException(
                    $"row[col op val] found no usable table on {scope}: a header '{h.name}' exists at " +
                    $"{h.sheet}!{h.cellRef}, but the surrounding cells are not a recognizable table. " +
                    "Auto-detection needs at least 2 adjacent non-empty header columns and at least 1 data " +
                    "row; a blank-header column in the middle or a header-only block breaks it. Create the " +
                    $"table explicitly (add <file> /{h.sheet} --type table --prop ref=<A1:D10>), or address cells directly.")
                {
                    Code = "not_found",
                    Suggestion = $"add <file> /{h.sheet} --type table --prop ref=<range covering the header and data>",
                };
            throw BuildNoColumnException(
                $"row[col op val] found no table on {scope} with column(s) {cols}. " +
                "Column predicates resolve header names (or column letters) against a ListObject or a detected (header-row) table.",
                scopeTables, colConds);
        }
        if (candidates.Count > 1)
        {
            var where = string.Join(", ", candidates.Select(c => $"{c.sheetName}!{c.label}"));
            var sheets = candidates.Select(c => c.sheetName).Distinct().ToList();
            // A sheet scope only disambiguates when the tables are on DIFFERENT
            // sheets. When they share one sheet, that guidance is a dead end —
            // there is no table-scope selector, so point at the tables' own
            // ranges instead.
            var suggestion = sheets.Count > 1
                ? "Scope by sheet, e.g. /SheetName/row[...]."
                : $"Both tables are on '{sheets[0]}', so a sheet scope cannot disambiguate — " +
                  $"address the intended table's rows by its cell range directly (e.g. {sheets[0]}!{candidates[0].label}).";
            throw new Core.CliException(
                $"row[col op val] is ambiguous — column(s) exist in {candidates.Count} tables ({where}). {suggestion}")
                { Code = "invalid_selector" };
        }

        var cand = candidates[0];
        var results = new List<DocumentNode>();
        var sheetData = GetSheet(cand.part).GetFirstChild<SheetData>();
        if (sheetData == null) return results;
        var eval = new Core.FormulaEvaluator(sheetData, _doc.WorkbookPart);

        // Index the physically-stored rows in the data range once. FindCell is
        // a full-sheet scan, so calling it per (row × predicate column) made a
        // 10k-row table cost O(rows × cells) ≈ seconds per query; one pass
        // keeps the whole match O(cells). Sparse rows stay probed via the
        // r-loop below (absent <row> → empty value), same as before.
        var rowsByIdx = new Dictionary<uint, Row>();
        foreach (var rw in sheetData.Elements<Row>())
        {
            var ri = rw.RowIndex?.Value ?? 0u;
            if (ri >= (uint)cand.dataR1 && ri <= (uint)cand.dataR2) rowsByIdx[ri] = rw;
        }
        for (int r = cand.dataR1; r <= cand.dataR2; r++)
        {
            rowsByIdx.TryGetValue((uint)r, out var xmlRow);
            // Probe node carries each predicate column's cell value under its
            // key so AttributeFilter evaluates all operators with one engine.
            var probe = new DocumentNode { Type = "cell" };
            foreach (var keyGroup in colConds.GroupBy(c => c.Key))
            {
                var wantRef = $"{IndexToColumnName(cand.colAbsIndex[keyGroup.Key])}{r}";
                Cell? cell = null;
                if (xmlRow != null)
                    foreach (var cc in xmlRow.Elements<Cell>())
                        if (cc.CellReference?.Value?.Equals(wantRef, StringComparison.OrdinalIgnoreCase) == true)
                        { cell = cc; break; }
                if (cell == null) { probe.Format[keyGroup.Key] = ""; continue; }
                // Compare on the underlying stored value (0.5 / date serial) when
                // any condition on this column is relational or carries a numeric
                // literal; otherwise the formatted display, so equality against a
                // formatted literal ("50%", "2024-01-15", text) still matches.
                bool wantsRaw = keyGroup.Any(c =>
                    c.Op is AttributeFilter.FilterOp.GreaterThan or AttributeFilter.FilterOp.LessThan
                        or AttributeFilter.FilterOp.GreaterOrEqual or AttributeFilter.FilterOp.LessOrEqual
                    || double.TryParse(c.Value, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out _));
                probe.Format[keyGroup.Key] = wantsRaw
                    ? GetCellRawComparisonValue(cell, eval)
                    : GetCellDisplayValue(cell, eval);
            }
            if (!AttributeFilter.MatchesExpr(probe, expr)) continue;

            var rowNode = new DocumentNode
            {
                Path = $"/{cand.sheetName}/row[{r}]",
                Type = "row",
                Preview = r.ToString(),
            };
            // Carry each predicate column's value under its column key so the
            // row is self-describing AND the CLI post-filter (which re-applies
            // the same [col op val] conditions on top of these results) resolves
            // the key and re-confirms the match, instead of dropping every row
            // because "Salary" is absent from a plain row node's Format.
            foreach (var cond in colConds)
                rowNode.Format[cond.Key] = probe.Format[cond.Key];
            // Cross-handler `evaluated` protocol: when a probed column
            // surfaced the #OCLI_NOTEVAL! sentinel, flag the row so callers
            // read Format["evaluated"] instead of string-sniffing the
            // sentinel out of the column value.
            if (colConds.Any(c => (probe.Format[c.Key]?.ToString() ?? "") == "#OCLI_NOTEVAL!"))
                rowNode.Format["evaluated"] = false;
            // Trace which table bound the predicate. source=detected flags a
            // header-sniff (stable=false) match so the caller knows the column
            // resolution was heuristic, mirroring DetectTables' own stable flag.
            rowNode.Format["matchedTable"] = cand.label;
            if (cand.source == "detected") rowNode.Format["tableSource"] = "detected";
            rowNode.ChildCount = xmlRow?.Elements<Cell>().Count() ?? 0;
            // Surface the row's own sparse-emit properties (hidden rows DO
            // match predicates — Excel-consistent), so a hit is self-describing
            // without a follow-up get row[N].
            if (xmlRow?.Hidden?.Value == true) rowNode.Format["hidden"] = true;
            results.Add(rowNode);
        }
        return results;
    }

    // Build a structured "no such column" error. The human message lists every
    // header per narrow table (<= 20 columns) or, for a wide table, only the
    // headers nearest (Damerau-Levenshtein) to the typo'd predicate keys plus a
    // pointer to `query "table" --json`. The SAME data is exposed as structured
    // CliException fields (Code/ValidValues/Suggestion/Help) so `--json` consumers
    // read machine fields instead of scraping the prose. When no table exists in
    // scope there is nothing to list — a bare not_found is returned.
    private const int ColumnListFullThreshold = 20;
    private const int ColumnSuggestTopK = 5;
    private static Core.CliException BuildNoColumnException(
        string baseMsg,
        List<(string label, List<string> cols)> scopeTables,
        List<AttributeFilter.Condition> colConds)
    {
        var tables = scopeTables
            .Select(t => (t.label, cols: t.cols.Where(c => !string.IsNullOrWhiteSpace(c)).ToList()))
            .Where(t => t.cols.Count > 0)
            .ToList();
        if (tables.Count == 0)
            return new Core.CliException(baseMsg) { Code = "not_found" };

        var wanted = colConds.Select(c => StripColPrefix(c.Key)).ToList();
        // Distinct headers across every in-scope table, first occurrence wins.
        var allCols = tables.SelectMany(t => t.cols)
            .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        bool anyWide = tables.Any(t => t.cols.Count > ColumnListFullThreshold);

        // Human-readable per-table suffix (text mode).
        var parts = new List<string>();
        foreach (var (label, cols) in tables)
        {
            if (cols.Count <= ColumnListFullThreshold)
                parts.Add($"table '{label}' has columns: {string.Join(", ", cols)}");
            else
                parts.Add($"table '{label}' has {cols.Count} columns; did you mean: " +
                    $"{string.Join(", ", NearestColumns(cols, wanted))}? (use 'query \"table\" --json' to list all)");
        }
        var message = baseMsg + " Available: " + string.Join("; ", parts) + ".";

        // Structured fields (--json). Narrow: enumerate all. Wide: rank nearest
        // and hand the caller the command that lists every column.
        if (!anyWide)
        {
            return new Core.CliException(message)
            {
                Code = "not_found",
                Suggestion = "Available columns: " + string.Join(", ", allCols),
                ValidValues = allCols.ToArray(),
            };
        }
        var nearest = NearestColumns(allCols, wanted);
        return new Core.CliException(message)
        {
            Code = "not_found",
            Suggestion = "Did you mean: " + string.Join(", ", nearest) + "?",
            ValidValues = nearest,
            Help = "Too many columns to list. Run 'query \"table\" --json' and read each " +
                   "table's `columns` field for the full list.",
        };
    }

    // Top-K headers ranked by smallest Damerau-Levenshtein distance to any key.
    private static string[] NearestColumns(List<string> cols, List<string> wanted)
        => cols
            .Select(col => (col, dist: wanted.Count == 0 ? int.MaxValue
                : wanted.Min(w => Core.EditDistance.Damerau(
                    w.ToLowerInvariant(), col.ToLowerInvariant()))))
            .OrderBy(x => x.dist)
            .Take(ColumnSuggestTopK)
            .Select(x => x.col)
            .ToArray();

    // Resolve every column predicate against a table's columns. Returns a
    // key→absolute-column-index map, or null if any predicate column is not in
    // this table (so the table is not a candidate).
    private static Dictionary<string, int>? ResolveColumns(
        List<string> colNames, int c1, int c2, List<AttributeFilter.Condition> colConds)
    {
        var resolved = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var cond in colConds)
        {
            if (!TryResolveTableColumn(cond.Key, colNames, c1, c2, out var absCol)) return null;
            resolved[cond.Key] = absCol;
        }
        return resolved;
    }

    // Resolve a predicate key to an ABSOLUTE column index within a table's
    // column span [c1..c2]. Header name (case-insensitive) wins over a column
    // letter so a header literally named "B" stays reachable by name.
    private static bool TryResolveTableColumn(string key, List<string> colNames, int c1, int c2, out int absCol)
    {
        absCol = 0;
        key = StripColPrefix(key);   // `col.Salary` → resolve against `Salary`
        var nameIdx = colNames.FindIndex(n => n.Equals(key, StringComparison.OrdinalIgnoreCase));
        if (nameIdx >= 0)
        {
            // Duplicate header WITHIN one table (only possible on detected
            // tables — Excel de-dupes real ListObject column names). Silently
            // taking the first match mis-targets half the time; make the
            // caller pick the exact column by LETTER instead.
            var dupIdx = colNames.FindIndex(nameIdx + 1, n => n.Equals(key, StringComparison.OrdinalIgnoreCase));
            if (dupIdx >= 0)
                throw new Core.CliException(
                    $"Column '{key}' is ambiguous: this table has {colNames.Count(n => n.Equals(key, StringComparison.OrdinalIgnoreCase))} " +
                    $"columns with that header. Address it by column letter instead, e.g. '{IndexToColumnName(c1 + nameIdx)}' or '{IndexToColumnName(c1 + dupIdx)}'.")
                {
                    Code = "invalid_selector",
                    ValidValues = colNames.Select((n, i) => (n, i))
                        .Where(x => x.n.Equals(key, StringComparison.OrdinalIgnoreCase))
                        .Select(x => IndexToColumnName(c1 + x.i)).ToArray(),
                };
            absCol = c1 + nameIdx;
            return true;
        }
        if (Regex.IsMatch(key, @"^[A-Za-z]{1,3}$"))
        {
            var letterIdx = ColumnNameToIndex(key.ToUpperInvariant());
            if (letterIdx >= c1 && letterIdx <= c2) { absCol = letterIdx; return true; }
        }
        return false;
    }

    // Set-side counterpart of the read predicate: resolve a
    // `set /Sheet/row[N] --prop <colName>=<val>` to the concrete cell in that
    // row's column, so the "filter then edit" loop closes symmetrically with
    // `query row[col op val]`. Binds to the single table (ListObject, else a
    // detected header-row table) on THIS worksheet that both owns column `key`
    // and has row N among its DATA rows (header/totals rows excluded). Returns
    // the A1 cell ref (e.g. "C5"). Throws on cross-table ambiguity; returns
    // false when nothing binds — the caller then reports the key as unsupported
    // rather than silently stamping it onto the <row> element (the old bug:
    // `set row[N] --prop 年龄=99` reported success but wrote an ignored XML attr).
    private bool TryResolveRowColumnCell(WorksheetPart worksheet, uint rowIdx, string key, out string cellRef)
    {
        cellRef = "";
        var listHits = new List<(string cell, string label)>();
        var detHits = new List<(string cell, string label)>();
        var realRanges = new List<(int c1, int r1, int c2, int r2)>();

        foreach (var tdp in worksheet.TableDefinitionParts)
        {
            var tbl = tdp.Table;
            if (tbl?.Reference?.Value == null) continue;
            if (!TryParseRange(tbl.Reference.Value, out var rng)) continue;
            realRanges.Add(rng);
            bool headerRow = (tbl.HeaderRowCount?.Value ?? 1) != 0;
            bool totalRow = (tbl.TotalsRowCount?.Value ?? 0) > 0 || (tbl.TotalsRowShown?.Value ?? false);
            if (rowIdx < rng.r1 + (headerRow ? 1 : 0) || rowIdx > rng.r2 - (totalRow ? 1 : 0)) continue;
            var colNames = tbl.GetFirstChild<TableColumns>()?.Elements<TableColumn>()
                .Select(c => c.Name?.Value ?? "").ToList() ?? new List<string>();
            if (TryResolveTableColumn(key, colNames, rng.c1, rng.c2, out var absCol))
                listHits.Add(($"{IndexToColumnName(absCol)}{rowIdx}", tbl.Name?.Value ?? "table"));
        }

        if (listHits.Count == 0)
        {
            var sheetName = GetWorksheets().FirstOrDefault(t => t.Part == worksheet).Name ?? "";
            foreach (var det in DetectTables(sheetName, worksheet, realRanges))
            {
                var colNames = DetectedTableColumns(det);
                var refStr = det.Format.TryGetValue("ref", out var rv) ? rv?.ToString() : null;
                if (!TryParseRange(refStr, out var frng)) continue;
                if (rowIdx < frng.r1 + 1 || rowIdx > frng.r2) continue;   // header sniff = 1 header row, no totals
                if (TryResolveTableColumn(key, colNames, frng.c1, frng.c2, out var absCol))
                    detHits.Add(($"{IndexToColumnName(absCol)}{rowIdx}", refStr!));
            }
        }

        var hits = listHits.Count > 0 ? listHits : detHits;
        if (hits.Count == 0) return false;
        var distinctCells = hits.Select(h => h.cell).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (distinctCells.Count > 1)
        {
            var where = string.Join(", ", hits.Select(h => h.label).Distinct());
            throw new Core.CliException(
                $"set row[{rowIdx}] --prop {StripColPrefix(key)}=… is ambiguous — column '{StripColPrefix(key)}' " +
                $"resolves in {distinctCells.Count} tables ({where}). Target the cell directly, e.g. {distinctCells[0]}.")
            { Code = "not_found" };
        }
        cellRef = distinctCells[0];
        return true;
    }

    // Error-path only: find a cell in scope whose text equals one of the
    // predicate column names. Lets the "no table" error distinguish "the header
    // exists but isn't part of a detectable table" from "the column name is
    // wrong", so a blank-header gap column or a header-only block gets accurate
    // guidance instead of a misleading typo hint. Scans raw cells (cheap, only
    // runs when a predicate already failed to bind).
    private (string sheet, string cellRef, string name)? FindHeaderLikeCell(
        string? sheetFilter, List<AttributeFilter.Condition> colConds)
    {
        var wanted = colConds.Select(c => StripColPrefix(c.Key)).ToList();
        foreach (var (sheetName, part) in GetWorksheets())
        {
            if (sheetFilter != null && !sheetName.Equals(sheetFilter, StringComparison.OrdinalIgnoreCase))
                continue;
            var sd = GetSheet(part).GetFirstChild<SheetData>();
            if (sd == null) continue;
            foreach (var row in sd.Elements<Row>())
                foreach (var cell in row.Elements<Cell>())
                {
                    var txt = GetCellDisplayValue(cell);
                    var w = wanted.FirstOrDefault(x => x.Equals(txt, StringComparison.OrdinalIgnoreCase));
                    if (w != null) return (sheetName, cell.CellReference?.Value ?? "?", w);
                }
        }
        return null;
    }

    // Human hint for a `set /Sheet/row[N] --prop <key>=…` whose key named
    // neither a row attribute nor a resolvable table column. Lists the columns
    // of the table(s) covering row N so the caller sees the real header (`薪水`
    // → `薪资`), mirroring the query-side guidance. The returned string embeds a
    // "(...)" hint, which suppresses the generic did-you-mean suggester
    // downstream (CommandBuilder.FormatUnsupported). Returns null when no table
    // covers the row — then the caller reports the bare key.
    private string? DescribeRowColumnsHint(WorksheetPart worksheet, uint rowIdx, string key)
    {
        var bare = StripColPrefix(key);
        // Every table on the sheet with its data-row span and column names, so
        // the hint can distinguish "row is IN a table but the column is wrong"
        // from "row is OUTSIDE the table but the column exists" (an append).
        var tables = new List<(string refStr, int dataR1, int dataR2, int c1, List<string> cols)>();
        var realRanges = new List<(int c1, int r1, int c2, int r2)>();
        foreach (var tdp in worksheet.TableDefinitionParts)
        {
            var tbl = tdp.Table;
            if (tbl?.Reference?.Value == null) continue;
            if (!TryParseRange(tbl.Reference.Value, out var rng)) continue;
            realRanges.Add(rng);
            bool headerRow = (tbl.HeaderRowCount?.Value ?? 1) != 0;
            bool totalRow = (tbl.TotalsRowCount?.Value ?? 0) > 0 || (tbl.TotalsRowShown?.Value ?? false);
            var cols = tbl.GetFirstChild<TableColumns>()?.Elements<TableColumn>()
                .Select(c => c.Name?.Value ?? "").ToList() ?? new List<string>();
            tables.Add((tbl.Reference.Value, rng.r1 + (headerRow ? 1 : 0), rng.r2 - (totalRow ? 1 : 0), rng.c1, cols));
        }
        if (tables.Count == 0)
        {
            var sheetName = GetWorksheets().FirstOrDefault(t => t.Part == worksheet).Name ?? "";
            foreach (var det in DetectTables(sheetName, worksheet, realRanges))
            {
                var refStr = det.Format.TryGetValue("ref", out var rv) ? rv?.ToString() : null;
                if (!TryParseRange(refStr, out var frng)) continue;
                tables.Add((refStr!, frng.r1 + 1, frng.r2, frng.c1, DetectedTableColumns(det)));
            }
        }

        static List<string> Clean(IEnumerable<string> cs) =>
            cs.Where(c => !string.IsNullOrWhiteSpace(c)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

        // 1. A table covers this row → the column name is simply wrong.
        var covering = tables.Where(t => rowIdx >= t.dataR1 && rowIdx <= t.dataR2).ToList();
        if (covering.Count > 0)
        {
            var cols = Clean(covering.SelectMany(t => t.cols));
            if (cols.Count == 0) return null;
            string shown = cols.Count <= ColumnListFullThreshold
                ? string.Join(", ", cols)
                : $"{string.Join(", ", NearestColumns(cols, new List<string> { bare }))} (of {cols.Count})";
            return $"{key} (no such column; available: {shown})";
        }

        // 2. Row is just BELOW a table whose columns include the key — the user
        // is trying to append a row by column name. Explain the boundary and the
        // append path instead of mis-suggesting a raw row attribute. Restricted
        // to rows past the data (not the header/above), where "append" is apt.
        var owning = tables.FirstOrDefault(t =>
            rowIdx > t.dataR2 && t.cols.Any(c => c.Equals(bare, StringComparison.OrdinalIgnoreCase)));
        if (owning.cols != null)
        {
            var letter = IndexToColumnName(owning.c1 +
                owning.cols.FindIndex(c => c.Equals(bare, StringComparison.OrdinalIgnoreCase)));
            return $"{key}: row {rowIdx} is outside table {owning.refStr}; '{bare}' is a column of that " +
                   $"table but column-name set only writes rows already in it. Append by writing cells by " +
                   $"address (e.g. {letter}{rowIdx}), then extend the table's range to include the new row";
        }
        return null;
    }

    // True when an in-scope table (ListObject or detected) has a column whose
    // name equals `key`. Used to flag a bare `row[key op val]` whose key also
    // names a row PROPERTY (height/hidden/...): rather than silently choosing the
    // property and shadowing the column, the caller errors and points to the
    // `col.key` / `@key` escapes. Only ever called for the ≤5 attribute names, so
    // the table scan is negligible.
    private bool RowKeyCollidesWithColumn(string key, string? sheetFilter)
    {
        foreach (var (sheetName, worksheetPart) in GetWorksheets())
        {
            if (sheetFilter != null && !sheetName.Equals(sheetFilter, StringComparison.OrdinalIgnoreCase))
                continue;

            var realRanges = new List<(int c1, int r1, int c2, int r2)>();
            foreach (var tdp in worksheetPart.TableDefinitionParts)
            {
                var tbl = tdp.Table;
                if (tbl?.Reference?.Value == null) continue;
                if (TryParseRange(tbl.Reference.Value, out var rng)) realRanges.Add(rng);
                var colNames = tbl.GetFirstChild<TableColumns>()?.Elements<TableColumn>()
                    .Select(c => c.Name?.Value ?? "") ?? Enumerable.Empty<string>();
                if (colNames.Any(n => n.Equals(key, StringComparison.OrdinalIgnoreCase)))
                    return true;
            }
            foreach (var det in DetectTables(sheetName, worksheetPart, realRanges))
            {
                var colNames = DetectedTableColumns(det);
                if (colNames.Any(n => n.Trim().Equals(key, StringComparison.OrdinalIgnoreCase)))
                    return true;
            }
        }
        return false;
    }
}
