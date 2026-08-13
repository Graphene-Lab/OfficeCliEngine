// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using Vml = DocumentFormat.OpenXml.Vml;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{

    // CONSISTENCY(style-dual-key): resolve a style display name to its
    // OOXML styleId by scanning the styles part. Returns null when no
    // matching style is found, letting callers fall back to using the
    // value verbatim (lenient input). Used by paragraph-level Set on
    // styleName so users can write back the canonical readback key.
    private string? ResolveStyleIdFromName(string displayName)
    {
        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles == null || string.IsNullOrEmpty(displayName)) return null;
        var match = stylesPart.Styles.Elements<Style>()
            .FirstOrDefault(s => string.Equals(s.StyleName?.Val?.Value, displayName, StringComparison.Ordinal));
        return match?.StyleId?.Value;
    }

    /// <summary>
    /// Returns true if a style with the given styleId exists in the Styles part.
    /// "Normal" is implicit in OOXML and considered to exist even when the
    /// blank-document StyleDefinitionsPart is empty/absent — matches Word's
    /// own behaviour where every doc has Normal as the default paragraph style.
    /// </summary>
    internal bool StyleIdExists(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId)) return false;
        if (string.Equals(styleId, "Normal", StringComparison.Ordinal)) return true;
        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles == null) return false;
        return stylesPart.Styles.Elements<Style>()
            .Any(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.Ordinal));
    }

    private string GetStyleName(Paragraph para)
    {
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return "Normal";

        // Try to resolve display name from styles part
        var stylesPart = _doc.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles != null)
        {
            var style = stylesPart.Styles.Elements<Style>()
                .FirstOrDefault(s => s.StyleId?.Value == styleId);
            if (style?.StyleName?.Val?.Value != null)
                return style.StyleName.Val.Value;
        }

        return styleId;
    }

    private static int GetHeadingLevel(string styleName)
    {
        // Heading 1, Heading 2, heading1, 标题 1, etc.
        foreach (var ch in styleName)
        {
            if (char.IsDigit(ch))
                return ch - '0';
        }
        if (styleName == "Title") return 0;
        if (styleName == "Subtitle") return 1;
        return 1;
    }

    // CONSISTENCY(outline-resolution): build a styleId -> outline level map
    // from each style's own w:outlineLvl. OOXML §17.3.1.20: w:val is
    // ST_DecimalNumber 0-8; surface it as a 1-based level (val + 1). Heading
    // styles carry this, so a paragraph can resolve to an outline level even
    // when its styleId is numeric or localized and its display name contains
    // no recognizable "Heading"/"标题" token.
    private Dictionary<string, int> BuildStyleOutlineLevels()
    {
        var map = new Dictionary<string, int>(StringComparer.Ordinal);
        var styles = _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles == null) return map;
        foreach (var s in styles.Elements<Style>())
        {
            var id = s.StyleId?.Value;
            if (string.IsNullOrEmpty(id)) continue;
            var lvl = s.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
            if (lvl.HasValue && lvl.Value >= 0 && lvl.Value <= 8)
                map[id] = lvl.Value + 1;
        }
        return map;
    }

    // Resolve a paragraph's outline level (0 = Title, 1-9 = outline depth),
    // or -1 when the paragraph is not part of the document outline. Signals,
    // in priority order, mirror TOC generation so `view outline` and TOC agree:
    //   1. direct paragraph w:outlineLvl (OOXML §17.3.1.20, val 0-8 -> level 1-9)
    //   2. the paragraph style's own w:outlineLvl (via BuildStyleOutlineLevels)
    //   3. style display name (Heading N / 标题 N / Title / Subtitle)
    private int GetParagraphOutlineLevel(Paragraph para, Dictionary<string, int> styleLevels, out string styleName)
    {
        styleName = GetStyleName(para);

        var direct = para.ParagraphProperties?.OutlineLevel?.Val?.Value;
        if (direct.HasValue && direct.Value >= 0 && direct.Value <= 8)
            return direct.Value + 1;

        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId != null && styleLevels.TryGetValue(styleId, out var sl))
            return sl;

        if (styleName.Contains("Heading") || styleName.Contains("标题")
            || styleName.StartsWith("heading", StringComparison.OrdinalIgnoreCase))
            return GetHeadingLevel(styleName);
        if (styleName == "Title") return 0;
        if (styleName == "Subtitle") return 1;
        return -1;
    }

    private static bool IsNormalStyle(string styleName)
    {
        return styleName is "Normal" or "正文" or "Body Text" or "Body" or "a"
            || styleName.StartsWith("Normal");
    }
}
