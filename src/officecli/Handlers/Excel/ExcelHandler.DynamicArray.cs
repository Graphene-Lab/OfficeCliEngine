// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Handlers;

/// <summary>
/// Dynamic-array (spill) writeback. A post-2016 dynamic-array formula
/// (SEQUENCE/FILTER/SORT/UNIQUE/MAP/…) only spills in Excel 365 when its anchor
/// cell carries BOTH the array CellFormula (<c>t="array"</c>) AND a
/// <c>cm="1"</c> cell-metadata index pointing at an <c>XLDAPR</c> dynamic-array
/// record in <c>xl/metadata.xml</c>. Without the metadata record Excel treats
/// <c>t="array"</c> as a legacy CSE array locked to the single anchor cell and
/// does NOT spill.
///
/// We write the anchor only — NOT the spilled "ghost" cells. Empirically (real
/// Excel for Mac) Excel recomputes the spill extent from the formula on open and
/// fills the region itself, so the anchor + metadata is sufficient and the
/// stored <c>ref</c> need not match the true extent. Leaving the ghost cells to
/// Excel keeps officecli out of the spill-region lifecycle entirely: no risk of
/// overwriting a user cell, no stale ghosts to garbage-collect, and Excel owns
/// <c>#SPILL!</c> conflict detection.
/// </summary>
public partial class ExcelHandler
{
    // The XLDAPR cell-metadata record, captured verbatim from a real Excel 365
    // dynamic-array workbook. Static for every workbook — a single record that
    // every dynamic-array anchor's cm="1" points at. (No XML declaration: the
    // typed Metadata element constructor expects a bare root element.)
    private const string DynamicArrayMetadataXml =
        "<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" " +
        "xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\">" +
        "<metadataTypes count=\"1\">" +
        "<metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" " +
        "pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" " +
        "clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\"/>" +
        "</metadataTypes>" +
        "<futureMetadata name=\"XLDAPR\" count=\"1\"><bk><extLst>" +
        "<ext uri=\"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}\">" +
        "<xda:dynamicArrayProperties fDynamic=\"1\" fCollapsed=\"0\"/>" +
        "</ext></extLst></bk></futureMetadata>" +
        "<cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata>" +
        "</metadata>";

    /// <summary>
    /// Make <paramref name="cell"/> a dynamic-array anchor: ensure the workbook
    /// carries the XLDAPR cell-metadata part and point the cell's <c>cm</c> at
    /// the dynamic-array record. Idempotent — the part is created once per
    /// workbook and reused by every subsequent spill formula.
    /// </summary>
    private void EnsureDynamicArrayMetadata(Cell cell)
    {
        var wbPart = _doc.WorkbookPart;
        if (wbPart == null) return;

        // Our metadata part holds exactly one cellMetadata record (the XLDAPR
        // dynamic-array record), so its 1-based index is always 1. The common
        // case is officecli writing one or more spill formulas into a file with
        // no prior cell metadata; we create the part once and every anchor
        // reuses index 1. (A workbook that already carries a foreign
        // CellMetadataPart is not produced by officecli; we reuse index 1 there
        // too — a mismatched pointer at worst suppresses the spill, never
        // corrupts the file.)
        if (wbPart.GetPartsOfType<CellMetadataPart>().FirstOrDefault() == null)
        {
            var part = wbPart.AddNewPart<CellMetadataPart>();
            part.Metadata = new Metadata(DynamicArrayMetadataXml);
        }
        cell.CellMetaIndex = 1U;
    }
}
