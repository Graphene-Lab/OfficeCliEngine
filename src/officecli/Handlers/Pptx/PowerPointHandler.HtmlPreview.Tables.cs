// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using OfficeCli.Core.TableStyles;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler
{
    // ==================== Table Rendering ====================

    private static void RenderTable(StringBuilder sb, GraphicFrame gf, Dictionary<string, string> themeColors, string? dataPath = null,
        (long x, long y, long cx, long cy)? overridePos = null, OpenXmlPart? part = null)
    {
        var dataPathAttr = string.IsNullOrEmpty(dataPath) ? "" : $" data-path=\"{HtmlEncode(dataPath)}\"";
        var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
        if (table == null) return;

        var offset = gf.Transform?.Offset;
        var extents = gf.Transform?.Extents;
        if (offset == null || extents == null) return;

        // R14-2: when nested in a group, the caller re-projects position/size into
        // the group's child coordinate system (CalcGroupChildPos).
        var x = overridePos?.x ?? offset.X?.Value ?? 0;
        var y = overridePos?.y ?? offset.Y?.Value ?? 0;
        var cx = overridePos?.cx ?? extents.Cx?.Value ?? 0;
        var cy = overridePos?.cy ?? extents.Cy?.Value ?? 0;

        // PowerPoint stores the graphicFrame's declared layout height in <p:xfrm>,
        // but tables auto-grow vertically to fit explicit row heights — declared cy
        // can underreport actual rendered height. With overflow:hidden on the
        // container, this clips trailing rows (slide 6 of test-samples/07.pptx
        // declared 72pt for a 5×30.2pt = 151pt table). Honor the larger of the
        // two so all rows render.
        var rowHeightSum = table.Elements<Drawing.TableRow>().Sum(r => r.Height?.Value ?? 0);
        if (rowHeightSum > cy) cy = rowHeightSum;

        // Same idea on the horizontal axis. <a:gridCol w="…"> stores absolute
        // EMU per column; PowerPoint renders the table at Σ gridCol.w and lets
        // it overflow the graphicFrame's <p:ext cx> (e.g. after `add column`
        // appends a new col, cx is unchanged while the grid grows). The frame
        // does not clip — the slide canvas does. Use Σ gridCol.w as the true
        // table width; fall back to cx only when <a:tblGrid> is absent.
        var gridCols = table.TableGrid?.Elements<Drawing.GridColumn>().ToList();
        long gridWidthSum = gridCols?.Sum(gc => gc.Width?.Value ?? 0) ?? 0;
        var tableWidthEmu = gridWidthSum > 0 ? gridWidthSum : cx;

        // Detect table style + banding flags. All cell-level styling
        // (fill, text color, borders) is now resolved through
        // Core/TableStyles/TableStyleResolver — no local catalogue lives in
        // this file. Unknown style ids resolve to null and the cell falls
        // back to "no fill / no border" (correct for un-styled tables).
        var tblPr = table.GetFirstChild<Drawing.TableProperties>();
        var tableStyleId = tblPr?.GetFirstChild<Drawing.TableStyleId>()?.InnerText;
        bool hasFirstRow = tblPr?.FirstRow?.Value == true;
        bool hasBandRow = tblPr?.BandRow?.Value == true;
        bool hasLastRow = tblPr?.LastRow?.Value == true;
        bool hasFirstCol = tblPr?.FirstColumn?.Value == true;
        bool hasLastCol = tblPr?.LastColumn?.Value == true;
        bool hasBandCol = tblPr?.BandColumn?.Value == true;
        int totalRows = table.Elements<Drawing.TableRow>().Count();
        int totalCols = gridCols?.Count ?? 0;

        // Whole-table background fill (<a:tblPr><a:solidFill|gradFill>). PowerPoint
        // renders this as the table-container background — visible through cell margins
        // and any noFill cells (Format Table -> Fill). The renderer previously read
        // tblPr only for the style id + banding flags and dropped this fill entirely.
        var tblBgCss = "";
        var tblSolid = tblPr?.GetFirstChild<Drawing.SolidFill>();
        var tblGrad = tblPr?.GetFirstChild<Drawing.GradientFill>();
        if (tblSolid != null)
        {
            var bg = ResolveFillColor(tblSolid, themeColors);
            if (bg != null) tblBgCss = $";background:{bg}";
        }
        else if (tblGrad != null)
        {
            var gradCss = GradientToCss(tblGrad, themeColors);
            if (!string.IsNullOrEmpty(gradCss)) tblBgCss = $";background:{gradCss}";
        }

        // Right-to-left table (<a:tblPr rtl="1">): PowerPoint reverses the visual column
        // order (first column on the right). CSS direction:rtl on the <table> reverses the
        // <td> order; cells then carry direction:ltr so their text stays laid out normally
        // (matches PowerPoint, which reverses columns but keeps Latin cell text unchanged).
        var tableRtl = tblPr?.RightToLeft?.Value == true;
        var tableDirCss = tableRtl ? ";direction:rtl" : "";

        sb.AppendLine($"    <div class=\"table-container\"{dataPathAttr} style=\"left:{Units.EmuToPt(x)}pt;top:{Units.EmuToPt(y)}pt;width:{Units.EmuToPt(tableWidthEmu)}pt;height:{Units.EmuToPt(cy)}pt{tblBgCss}\">");
        sb.AppendLine($"      <table class=\"slide-table\" style=\"{(tableRtl ? "direction:rtl" : "")}\">");

        // Column widths — emit absolute pt per <a:gridCol w>, not percentages.
        // table-layout:fixed + width:100% on .slide-table then preserves these
        // widths (container width == Σ gridCol.w so they add up exactly).
        if (gridCols != null && gridCols.Count > 0)
        {
            sb.Append("        <colgroup>");
            foreach (var gc in gridCols)
            {
                var w = gc.Width?.Value ?? 0;
                if (w > 0)
                    sb.Append($"<col style=\"width:{Units.EmuToPt(w):0.##}pt\">");
                else
                    sb.Append($"<col style=\"width:{(100.0 / gridCols.Count):0.##}%\">");
            }
            sb.AppendLine("</colgroup>");
        }

        int rowIndex = 0;
        foreach (var row in table.Elements<Drawing.TableRow>())
        {
            // Honor explicit per-row height from <a:tr h="EMU">. Without this,
            // every row collapses to equal height (HTML table default), losing
            // the per-row sizing users set via `set tr[N] --prop height=`.
            var rowH = row.Height?.Value ?? 0;
            var rowStyle = rowH > 0 ? $" style=\"height:{Units.EmuToPt(rowH):0.##}pt\"" : "";
            sb.AppendLine($"        <tr{rowStyle}>");
            int colIndex = 0;  // Tracked for the new per-cell TableStyleResolver below.
            // 1-based TableCell element position — matches officecli's cell[C]
            // numbering (Set.Table indexes `row.Elements<TableCell>()`), so an
            // editor can map a clicked <td> to /row[R]/cell[C] and set its text.
            int cellElemIndex = 0;
            bool isHeaderRow = hasFirstRow && rowIndex == 0;
            bool isBandedOdd = hasBandRow && (!hasFirstRow ? rowIndex % 2 == 0 : rowIndex > 0 && (rowIndex - 1) % 2 == 0);

            foreach (var cell in row.Elements<Drawing.TableCell>())
            {
                cellElemIndex++;
                var cellStyles = new List<string>();
                // In an RTL table the <table> carries direction:rtl to reverse the column
                // order; keep each cell's content laid out LTR so text/alignment is
                // unchanged (PowerPoint reverses columns, not in-cell text flow).
                if (tableRtl) cellStyles.Add("direction:ltr");

                // Cell fill
                var tcPr = cell.TableCellProperties ?? cell.GetFirstChild<Drawing.TableCellProperties>();
                var cellSolid = tcPr?.GetFirstChild<Drawing.SolidFill>();
                var cellColor = ResolveFillColor(cellSolid, themeColors);
                bool hasExplicitFill = cellColor != null;
                if (cellColor != null)
                    cellStyles.Add($"background:{cellColor}");

                var cellGrad = tcPr?.GetFirstChild<Drawing.GradientFill>();
                if (cellGrad != null)
                {
                    cellStyles.Add($"background:{GradientToCss(cellGrad, themeColors)}");
                    hasExplicitFill = true;
                }

                // Pattern fill (<a:tcPr><a:pattFill>) — schema-valid on a table cell
                // (CT_TableCellProperties) just like solid/grad/blip/noFill. The shape
                // and slide-background fill paths already approximate it via
                // PatternFillToCss (repeating-linear-gradient); the cell path dropped
                // it, so a pattern-filled cell rendered with no fill. PatternFillToCss
                // returns a full "background:..." declaration.
                var cellPatt = tcPr?.GetFirstChild<Drawing.PatternFill>();
                if (cellPatt != null)
                {
                    cellStyles.Add(PatternFillToCss(cellPatt, themeColors));
                    hasExplicitFill = true;
                }

                // Picture fill (<a:tcPr><a:blipFill>): resolve the blip r:embed
                // against the part the table lives in (same rels as the slide).
                // PowerPoint stretches the image to fill the cell when <a:stretch>
                // is present (the common authoring case); mirror with
                // background-size:100% 100%. Without a part to resolve the rel,
                // skip silently (no regression for blip-less cells).
                var cellBlip = tcPr?.GetFirstChild<Drawing.BlipFill>();
                if (cellBlip != null && part != null)
                {
                    var dataUri = BlipToDataUri(cellBlip, part);
                    if (dataUri != null)
                    {
                        var sizeCss = cellBlip.GetFirstChild<Drawing.Stretch>() != null
                            ? "100% 100%" : "cover";
                        cellStyles.Add($"background-image:url('{dataUri}')");
                        cellStyles.Add($"background-size:{sizeCss}");
                        cellStyles.Add("background-position:center");
                        hasExplicitFill = true;
                    }
                }

                // Explicit <a:noFill/> is an intentional "transparent" declaration
                // that overrides any table-style band/header fill. PowerPoint shows
                // the slide background through such a cell. Mark hasExplicitFill so
                // the TableStyleResolver fill below is suppressed; emit no
                // background: declaration (transparent).
                if (tcPr?.GetFirstChild<Drawing.NoFill>() != null)
                    hasExplicitFill = true;

                // Resolve fill / text color / borders for this cell through
                // the Core/TableStyles catalogue. Returns null for unknown
                // style ids (custom styles, no style at all); in that case
                // the cell renders with no style-provided fill or borders,
                // which matches OOXML "no style" semantics.
                var resolved = TableStyleResolver.Resolve(
                    tableStyleId,
                    new CellPosition(
                        RowIndex: rowIndex, ColIndex: colIndex,
                        RowCount: totalRows, ColCount: totalCols,
                        HasFirstRow: hasFirstRow, HasLastRow: hasLastRow,
                        HasFirstCol: hasFirstCol, HasLastCol: hasLastCol,
                        HasBandedRows: hasBandRow, HasBandedCols: hasBandCol),
                    themeColors);

                // An explicit cell fill overrides only the table-style BACKGROUND.
                // The style's text color must still apply (PowerPoint keeps the
                // header row's white text when you override just the cell fill —
                // verified against real PowerPoint). Splitting the gate: Fill stays
                // suppressed by hasExplicitFill, TextColor does not. An explicit run
                // color is emitted on the run span and still wins by specificity.
                if (resolved != null)
                {
                    if (!hasExplicitFill && resolved.Fill != null)
                        cellStyles.Add($"background:{resolved.Fill}");
                    if (resolved.TextColor != null)
                        cellStyles.Add($"color:{resolved.TextColor}");
                }

                // Vertical alignment
                if (tcPr?.Anchor?.HasValue == true)
                {
                    var va = tcPr.Anchor.InnerText switch
                    {
                        "ctr" => "middle",
                        "b" => "bottom",
                        _ => "top"
                    };
                    cellStyles.Add($"vertical-align:{va}");
                }

                // Cell text direction (a:tcPr vert="…"): rotate text via CSS
                // writing-mode. Mirrors the Word handler's tcDir → writing-mode
                // mapping (WordHandler.HtmlPreview.Css.cs). PowerPoint stores
                // the direction on the tcPr "vert" attribute:
                //   vert    → top-to-bottom (vertical-rl)
                //   vert270 → bottom-to-top (vertical-rl + 180° rotate)
                //   eaVert  → East-Asian top-to-bottom, upright glyphs
                //   wordArtVert → stacked upright (approximate with upright)
                var vertDir = tcPr?.Vertical?.HasValue == true ? tcPr.Vertical.InnerText : null;
                string? cellTextTransform = null;
                if (vertDir != null)
                {
                    var wm = vertDir switch
                    {
                        "vert" => "vertical-rl",
                        "vert270" => "vertical-rl",
                        "eaVert" => "vertical-rl;text-orientation:upright",
                        "wordArtVert" or "wordArtVertRtl" => "vertical-rl;text-orientation:upright",
                        _ => null,
                    };
                    if (wm != null) cellStyles.Add($"writing-mode:{wm}");
                    // vert270 = 90° CCW (bottom-to-top): vertical-rl gives 90° CW, so flip
                    // 180°. CSS `transform` is IGNORED on a display:table-cell <td>, so the
                    // rotation must go on an inner wrapper around the cell content, not the
                    // <td>'s own style (which is why vert270 previously rendered like vert).
                    if (vertDir == "vert270") cellTextTransform = "rotate(180deg)";
                }

                // Cell text formatting
                var firstRun = cell.Descendants<Drawing.Run>().FirstOrDefault();
                if (firstRun?.RunProperties != null)
                {
                    var rp = firstRun.RunProperties;
                    if (rp.FontSize?.HasValue == true)
                        cellStyles.Add($"font-size:{rp.FontSize.Value / 100.0:0.##}pt");
                    // else: inherit from table style / slideMaster (no hardcoded default)
                    // (bold decided below — unified with the table-style emphasis bold)
                    // Italic / underline / strike — mirror RenderRun so cell text
                    // formatting matches shape text. Previously only bold was read,
                    // silently dropping <a:rPr i="1"/u="…"/strike="…"> on table cells.
                    if (rp.Italic?.Value == true)
                        cellStyles.Add("font-style:italic");
                    var decorations = new List<string>();
                    if (rp.Underline?.HasValue == true && rp.Underline.Value != Drawing.TextUnderlineValues.None)
                        decorations.Add("underline");
                    if (rp.Strike?.HasValue == true && rp.Strike.Value != Drawing.TextStrikeValues.NoStrike)
                        decorations.Add("line-through");
                    if (decorations.Count > 0)
                        cellStyles.Add($"text-decoration:{string.Join(" ", decorations)}");
                    var fontVal = rp.GetFirstChild<Drawing.LatinFont>()?.Typeface?.Value
                        ?? rp.GetFirstChild<Drawing.EastAsianFont>()?.Typeface?.Value;
                    if (fontVal != null && !fontVal.StartsWith("+", StringComparison.Ordinal))
                        cellStyles.Add(CssFontFamilyWithFallback(fontVal));
                    var runColor = ResolveFillColor(rp.GetFirstChild<Drawing.SolidFill>(), themeColors);
                    if (runColor != null)
                        cellStyles.Add($"color:{runColor}");
                }

                // Bold: an explicit run bold (b="1"/b="0") wins; otherwise the
                // table style's emphasis-band bold (header/total/first-col/last-col)
                // applies. Built-in PowerPoint styles render those bands bold, which
                // the renderer previously dropped (only explicit run bold was read).
                bool? explicitBold = firstRun?.RunProperties?.Bold?.Value;
                if (explicitBold ?? (resolved?.Bold == true))
                    cellStyles.Add("font-weight:bold");

                // Cell borders (per-edge). Priority cascade:
                //   1. Explicit <a:lnL/R/T/B> on this cell (per-cell override)
                //   2. TableStyleResolver output (built-in style catalogue)
                //   3. default thin-black border when the table has no resolvable
                //      style at all (tableStyleId null/unknown), else "none"
                // Explicit <a:lnL> with <a:noFill/> yields "none" via
                // TableBorderToCss and short-circuits cleanly. The resolver
                // computes per-cell borders based on position (outer vs.
                // inner edges) following the style's <a:tcBdr> region rules.
                // When no style resolves, native PowerPoint still applies the
                // deck's default table style (thin black grid); mirror that with
                // a 1pt solid #000000 fallback rather than rendering border-less.
                // CONSISTENCY(table-borders): Npt solid #color idiom.
                string FormatBorder(ResolvedBorder? rb)
                    => rb != null ? $"{Units.EmuToPt(rb.WidthEmu):0.##}pt {rb.Dash} {rb.Color}"
                       : (resolved == null ? "1pt solid #000000" : "none");
                var borderLeft = tcPr?.GetFirstChild<Drawing.LeftBorderLineProperties>();
                var borderRight = tcPr?.GetFirstChild<Drawing.RightBorderLineProperties>();
                var borderTop = tcPr?.GetFirstChild<Drawing.TopBorderLineProperties>();
                var borderBottom = tcPr?.GetFirstChild<Drawing.BottomBorderLineProperties>();
                var bl = TableBorderToCss(borderLeft, themeColors) ?? FormatBorder(resolved?.Left);
                var br = TableBorderToCss(borderRight, themeColors) ?? FormatBorder(resolved?.Right);
                var bt = TableBorderToCss(borderTop, themeColors) ?? FormatBorder(resolved?.Top);
                var bb = TableBorderToCss(borderBottom, themeColors) ?? FormatBorder(resolved?.Bottom);
                cellStyles.Add($"border-left:{bl}");
                cellStyles.Add($"border-right:{br}");
                cellStyles.Add($"border-top:{bt}");
                cellStyles.Add($"border-bottom:{bb}");

                // Diagonal borders (<a:lnTlToBr> / <a:lnBlToTr>) — HTML has no
                // native diagonal-border; emit an absolute-positioned inline
                // SVG overlay inside the <td>. The <td> becomes position:relative
                // only when diagonals are actually present to minimize CSS
                // regression surface.
                var borderTlBr = tcPr?.GetFirstChild<Drawing.TopLeftToBottomRightBorderLineProperties>();
                var borderBlTr = tcPr?.GetFirstChild<Drawing.BottomLeftToTopRightBorderLineProperties>();
                var tlBrCss = TableBorderToCss(borderTlBr, themeColors);
                var blTrCss = TableBorderToCss(borderBlTr, themeColors);
                bool hasDiag = (tlBrCss != null && tlBrCss != "none")
                            || (blTrCss != null && blTrCss != "none");
                if (hasDiag)
                    cellStyles.Add("position:relative");

                // Cell margins/padding
                var marL = tcPr?.LeftMargin?.Value;
                var marR = tcPr?.RightMargin?.Value;
                var marT = tcPr?.TopMargin?.Value;
                var marB = tcPr?.BottomMargin?.Value;
                // Always emit padding from OOXML defaults for absent attrs
                // (marL=marR=91440 EMU=7.2pt, marT=marB=45720 EMU=3.6pt). The
                // preview.css fallback (4px 6px) under-pads L/R to ~60% of the
                // correct value, so an all-default cell rendered text too far left.
                var pT = Units.EmuToPt(marT ?? 45720);
                var pR = Units.EmuToPt(marR ?? 91440);
                var pB = Units.EmuToPt(marB ?? 45720);
                var pL = Units.EmuToPt(marL ?? 91440);
                cellStyles.Add($"padding:{pT}pt {pR}pt {pB}pt {pL}pt");

                // Paragraph alignment is emitted PER-PARAGRAPH by RenderTextBody
                // (each <div class="para"> carries its own text-align). We must NOT
                // set a cell-level text-align from the first paragraph: CSS
                // text-align inherits, so a right-aligned first paragraph would
                // bleed onto subsequent paragraphs that have no explicit alignment
                // (PowerPoint renders those left/default). Leave the <td> without a
                // text-align so unaligned paragraphs fall back to the default.

                // Render the cell's paragraphs through the same path shape text
                // bodies use (RenderTextBody → one <div class="para"> per
                // paragraph), so a multi-paragraph cell shows each paragraph on
                // its own line instead of being flattened to InnerText (which
                // concatenated "Line 1Line 2Line 3"). Per-paragraph alignment /
                // spacing / bullets carry over for free. The <td>-level font /
                // color / align styles above still apply and are inherited by
                // the para divs, so a single-paragraph cell is visually
                // unchanged. Empty text body falls back to "" (renders nothing).
                var cellBody = new StringBuilder();
                if (cell.TextBody != null)
                    // Forward `part` as placeholderPart so cell text resolves the deck's
                    // theme font (RenderTextBody's themeFontFallback) and hyperlink
                    // relationships — matching shape text. Without it, runs with no
                    // explicit a:latin emitted no font-family and rendered in the page
                    // default instead of the theme minor font.
                    RenderTextBody(cellBody, cell.TextBody, themeColors,
                        placeholderShape: null, placeholderPart: part, fontRefDefaultColor: null);
                var cellHtml = cellBody.ToString();
                var styleStr = cellStyles.Count > 0 ? $" style=\"{string.Join(";", cellStyles)}\"" : "";

                // Column/row span (GridSpan and RowSpan are on the TableCell, not TableCellProperties)
                var gridSpan = cell.GridSpan?.Value;
                var rowSpan = cell.RowSpan?.Value;
                var spanAttrs = "";
                if (gridSpan > 1) spanAttrs += $" colspan=\"{gridSpan}\"";
                if (rowSpan > 1) spanAttrs += $" rowspan=\"{rowSpan}\"";

                // Skip merged continuation cells. The colIndex advance for the
                // whole horizontal span is done ONCE on the anchor's `colIndex +=
                // gridSpan` at the end of the loop body; continuation cells must
                // therefore NOT advance colIndex again (doing so over-counted the
                // span by gridSpan-1 and shifted banding/firstCol/lastCol for every
                // cell after a horizontal merge). hMerge continuation cells just
                // `continue` without rendering their own <td>; the anchor already
                // accounted for their columns. vMerge continuation cells are a
                // single-column rowspan body, so they DO advance colIndex by 1.
                //
                // BUT a cell that is BOTH hMerge AND vMerge (the lower continuation
                // rows of a 2-D merge block, e.g. the bottom-right of a 2x2 merge)
                // lives in a row whose horizontal-span anchor is in a DIFFERENT row
                // (the merge's top row), so THIS row never sees an anchor `colIndex
                // += gridSpan` to account for it. Such a cell still occupies exactly
                // one grid column in its own row and must advance colIndex by 1,
                // mirroring the pure-vMerge body case. Failing to advance shifted
                // the column-band fill of every cell after the merge by one column.
                if (cell.HorizontalMerge?.Value == true)
                {
                    if (cell.VerticalMerge?.Value == true)
                        colIndex++;
                    continue;
                }
                if (cell.VerticalMerge?.Value == true)
                {
                    colIndex++;
                    continue;
                }

                var diagOverlay = "";
                if (hasDiag)
                {
                    var diagLines = new StringBuilder();
                    if (tlBrCss != null && tlBrCss != "none")
                    {
                        var (stroke, widthPt) = ParseBorderCssForSvg(tlBrCss);
                        diagLines.Append($"<line x1=\"0\" y1=\"0\" x2=\"100%\" y2=\"100%\" stroke=\"{stroke}\" stroke-width=\"{widthPt:0.##}pt\"/>");
                    }
                    if (blTrCss != null && blTrCss != "none")
                    {
                        var (stroke, widthPt) = ParseBorderCssForSvg(blTrCss);
                        diagLines.Append($"<line x1=\"0\" y1=\"100%\" x2=\"100%\" y2=\"0\" stroke=\"{stroke}\" stroke-width=\"{widthPt:0.##}pt\"/>");
                    }
                    diagOverlay = $"<svg class=\"cell-diag\" width=\"100%\" height=\"100%\" style=\"position:absolute;inset:0;pointer-events:none;overflow:visible\" preserveAspectRatio=\"none\">{diagLines}</svg>";
                }

                // vert270 rotation must wrap the content (transform is inert on the <td>).
                if (cellTextTransform != null)
                    cellHtml = $"<div style=\"transform:{cellTextTransform};display:inline-block\">{cellHtml}</div>";
                // A SEPARATE attribute (not data-path) so selection/drag still
                // treat the whole table as the unit — an editor uses this only to
                // map a double-clicked cell to /row[R]/cell[C] for inline text edit.
                var cellPathAttr = string.IsNullOrEmpty(dataPath) ? ""
                    : $" data-cell-path=\"{HtmlEncode($"{dataPath}/row[{rowIndex + 1}]/cell[{cellElemIndex}]")}\"";
                sb.AppendLine($"          <td{spanAttrs}{cellPathAttr}{styleStr}>{diagOverlay}{cellHtml}</td>");
                colIndex += Math.Max((int)(gridSpan ?? 1), 1);
            }
            sb.AppendLine("        </tr>");
            rowIndex++;
        }

        sb.AppendLine("      </table>");
        sb.AppendLine("    </div>");
    }

    /// <summary>
    /// Convert a table cell border line properties element to a CSS border value.
    /// Returns null if the border has NoFill or is absent.
    /// </summary>
    private static string? TableBorderToCss(OpenXmlCompositeElement? borderProps, Dictionary<string, string> themeColors)
    {
        if (borderProps == null) return null;
        if (borderProps.GetFirstChild<Drawing.NoFill>() != null) return "none";

        var solidFill = borderProps.GetFirstChild<Drawing.SolidFill>();
        var color = ResolveFillColor(solidFill, themeColors) ?? "#000000";

        // Width attribute is on the element itself (w attr in EMU).
        // An EXPLICIT w="0" means "no border" (PowerPoint renders no line on
        // that edge); only an ABSENT w falls back to the default thin line.
        double widthPt = 1.0;
        long? widthEmu = borderProps switch
        {
            Drawing.LeftBorderLineProperties lb when lb.Width?.HasValue == true => lb.Width.Value,
            Drawing.RightBorderLineProperties rb when rb.Width?.HasValue == true => rb.Width.Value,
            Drawing.TopBorderLineProperties tb when tb.Width?.HasValue == true => tb.Width.Value,
            Drawing.BottomBorderLineProperties bb when bb.Width?.HasValue == true => bb.Width.Value,
            Drawing.TopLeftToBottomRightBorderLineProperties tlbr when tlbr.Width?.HasValue == true => tlbr.Width.Value,
            Drawing.BottomLeftToTopRightBorderLineProperties bltr when bltr.Width?.HasValue == true => bltr.Width.Value,
            _ => null
        };
        if (widthEmu == 0) return "none";
        if (widthEmu.HasValue) widthPt = widthEmu.Value / EmuConverter.EmuPerPointF;

        if (widthPt < 0.5) widthPt = 0.5;

        var dash = borderProps.GetFirstChild<Drawing.PresetDash>();
        var style = "solid";
        if (dash?.Val?.HasValue == true)
        {
            // CONSISTENCY(dash-pattern): map mixed dash-dot patterns to "dashed" (CSS has no native dashDot).
            // Previously fell through to "solid", which silently dropped the dash pattern.
            style = dash.Val.InnerText switch
            {
                "dash" or "lgDash" or "sysDash" => "dashed",
                "dot" or "sysDot" => "dotted",
                "dashDot" or "lgDashDot" or "lgDashDotDot"
                    or "sysDashDot" or "sysDashDotDot" => "dashed",
                _ => "solid"
            };
        }
        else if (borderProps.GetFirstChild<Drawing.CustomDash>() != null)
        {
            // CONSISTENCY(dash-pattern): <a:custDash> (mutually exclusive with prstDash)
            // also makes the border dashed. A CSS border-style cannot express an
            // arbitrary dash array, so approximate to "dashed" — same compromise the
            // dashDot presets above take. Without this a custom-dashed border rendered
            // SOLID (the shape/connector path emits a real stroke-dasharray; a CSS <td>
            // border can't, so dashed is the faithful approximation).
            style = "dashed";
        }

        // CONSISTENCY(compound-line): a border <a:ln cmpd="dbl"/thickThin/thinThick/tri>
        // renders as parallel lines in PowerPoint. CSS border-style has only "double",
        // so map any non-single compound type to "double" (mirrors the shape outline
        // path, which emits border-style:double for cmpd != "sng"). Only applies to an
        // otherwise-solid border — a dashed compound border keeps the dash approximation.
        if (style == "solid")
        {
            var cmpd = borderProps switch
            {
                Drawing.LeftBorderLineProperties lb => lb.CompoundLineType?.InnerText,
                Drawing.RightBorderLineProperties rb => rb.CompoundLineType?.InnerText,
                Drawing.TopBorderLineProperties tb => tb.CompoundLineType?.InnerText,
                Drawing.BottomBorderLineProperties bb => bb.CompoundLineType?.InnerText,
                Drawing.TopLeftToBottomRightBorderLineProperties tlbr => tlbr.CompoundLineType?.InnerText,
                Drawing.BottomLeftToTopRightBorderLineProperties bltr => bltr.CompoundLineType?.InnerText,
                _ => null
            };
            if (!string.IsNullOrEmpty(cmpd) && cmpd != "sng")
                style = "double";
        }

        return $"{widthPt:0.##}pt {style} {color}";
    }

    /// <summary>
    /// Parse the "Npt style #color" shorthand produced by TableBorderToCss
    /// back into (stroke-color, stroke-width-in-pt) for SVG diagonal lines.
    /// Format is deterministic: "{w:0.##}pt {solid|dashed|dotted} {color}".
    /// </summary>
    private static (string stroke, double widthPt) ParseBorderCssForSvg(string css)
    {
        var parts = css.Split(' ', 3, StringSplitOptions.RemoveEmptyEntries);
        double widthPt = 1.0;
        string stroke = "#000000";
        if (parts.Length >= 1)
        {
            var w = parts[0];
            if (w.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
                w = w[..^2];
            double.TryParse(w, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out widthPt);
        }
        if (parts.Length >= 3)
            stroke = parts[2];
        return (stroke, widthPt);
    }

    // Per-cell style resolution (fill, text color, borders) moved to
    // Core/TableStyles/TableStyleResolver. This file now contains only
    // OOXML→HTML rendering glue; the built-in PowerPoint table-style
    // catalogue (11 family templates × 7 accent variants = 74 GUIDs) lives
    // under Core/TableStyles/ with one file per family. See
    // the Core/TableStyles notes and the PPTX handler conventions.
}
