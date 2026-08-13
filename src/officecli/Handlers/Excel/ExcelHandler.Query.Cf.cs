// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    /// <summary>
    /// Populate a conditionalFormatting DocumentNode from a single cfRule.
    /// Extracted from the cf[N] Get path so the dump serializer can reuse it
    /// per-rule: a &lt;conditionalFormatting&gt; element may hold multiple
    /// &lt;cfRule&gt; children (Excel stacks rules on one range), and emitting
    /// only rule[0] silently dropped rule[1..N] on dump/replay.
    /// </summary>
    internal void PopulateCfNodeFromRule(WorksheetPart worksheet, ConditionalFormattingRule rule, DocumentNode cfNode)
    {
        // Canonical CF type key. Normalized variants overwrite this below
        // (e.g. ConditionalFormatValues.Top10 -> "topN", Expression -> "formula").
        if (rule.Type?.Value != null)
            cfNode.Format["type"] = rule.Type.InnerText;

        // stopIfTrue applies to every CF rule type; surface it so the
        // dump→batch round-trip can re-emit the attribute. Only emit
        // when explicitly true (OOXML default is false).
        if (rule.StopIfTrue?.Value == true)
            cfNode.Format["stopIfTrue"] = true;

        // DataBar
        var dataBar = rule.GetFirstChild<DataBar>();
        if (dataBar != null)
        {
            cfNode.Format["type"] = "dataBar";
            var dbColor = dataBar.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>();
            if (dbColor?.Rgb?.Value != null)
                cfNode.Format["color"] = ParseHelpers.FormatHexColor(dbColor.Rgb.Value);
            else if (dbColor?.Theme?.Value != null)
                cfNode.Format["color"] = $"theme{dbColor.Theme.Value}";
            // ShowValue defaults to true; only emit when explicitly false on the OOXML
            if (dataBar.ShowValue?.Value == false) cfNode.Format["showValue"] = false;
            if (dataBar.MinLength?.Value is uint dbMinLen) cfNode.Format["minLength"] = dbMinLen;
            if (dataBar.MaxLength?.Value is uint dbMaxLen) cfNode.Format["maxLength"] = dbMaxLen;

            // x14 extension: direction, negativeColor, axisColor
            var dbExtList = rule.GetFirstChild<ConditionalFormattingRuleExtensionList>();
            if (dbExtList != null)
            {
                // Look up the matching x14:cfRule by id reference; fall back to scanning worksheet extLst
                var x14CfRule = FindMatchingX14DataBarRule(GetSheet(worksheet), dbExtList);
                var x14Db = x14CfRule?.GetFirstChild<X14.DataBar>();
                if (x14Db != null)
                {
                    if (x14Db.Direction?.HasValue == true)
                        cfNode.Format["direction"] = x14Db.Direction.InnerText;
                    var negCol = x14Db.GetFirstChild<X14.NegativeFillColor>();
                    if (negCol?.Rgb?.Value != null)
                        cfNode.Format["negativeColor"] = ParseHelpers.FormatHexColor(negCol.Rgb.Value);
                    var axCol = x14Db.GetFirstChild<X14.BarAxisColor>();
                    if (axCol?.Rgb?.Value != null)
                        cfNode.Format["axisColor"] = ParseHelpers.FormatHexColor(axCol.Rgb.Value);
                    // minLength/maxLength live on the x14 dataBar (Add writes
                    // them there, not on the 2007 DataBar). Surface both so
                    // the round-trip preserves the bar-length bounds.
                    if (x14Db.MinLength?.Value is uint x14MinLen)
                        cfNode.Format["minLength"] = x14MinLen;
                    if (x14Db.MaxLength?.Value is uint x14MaxLen)
                        cfNode.Format["maxLength"] = x14MaxLen;
                    // axisPosition (middle/none/automatic). automatic is the
                    // OOXML default; only surface an explicit non-automatic value.
                    if (x14Db.AxisPosition?.HasValue == true
                        && x14Db.AxisPosition.Value != X14.DataBarAxisPositionValues.Automatic)
                        cfNode.Format["axisPosition"] = x14Db.AxisPosition.InnerText;
                    // Explicit num-typed min/max cfvo bounds. Surface the
                    // literal value so autoMin/autoMax don't silently replace
                    // user-set numeric bounds on replay.
                    var x14Cfvos = x14Db.Elements<X14.ConditionalFormattingValueObject>().ToList();
                    if (x14Cfvos.Count >= 1
                        && x14Cfvos[0].Type?.Value == X14.ConditionalFormattingValueObjectTypeValues.Numeric)
                    {
                        var f = x14Cfvos[0].GetFirstChild<DocumentFormat.OpenXml.Office.Excel.Formula>();
                        if (!string.IsNullOrEmpty(f?.Text)) cfNode.Format["min"] = f!.Text;
                    }
                    if (x14Cfvos.Count >= 2
                        && x14Cfvos[1].Type?.Value == X14.ConditionalFormattingValueObjectTypeValues.Numeric)
                    {
                        var f = x14Cfvos[1].GetFirstChild<DocumentFormat.OpenXml.Office.Excel.Formula>();
                        if (!string.IsNullOrEmpty(f?.Text)) cfNode.Format["max"] = f!.Text;
                    }
                }
            }
        }

        // ColorScale
        var colorScale = rule.GetFirstChild<ColorScale>();
        if (colorScale != null)
        {
            cfNode.Format["type"] = "colorScale";
            var colors = colorScale.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().ToList();
            if (colors.Count >= 2)
            {
                var minRgb = colors[0].Rgb?.Value;
                var maxRgb = colors[^1].Rgb?.Value;
                if (!string.IsNullOrEmpty(minRgb))
                    cfNode.Format["minColor"] = ParseHelpers.FormatHexColor(minRgb);
                if (!string.IsNullOrEmpty(maxRgb))
                    cfNode.Format["maxColor"] = ParseHelpers.FormatHexColor(maxRgb);
                if (colors.Count >= 3)
                {
                    var midRgb = colors[1].Rgb?.Value;
                    if (!string.IsNullOrEmpty(midRgb))
                        cfNode.Format["midColor"] = ParseHelpers.FormatHexColor(midRgb);
                    // Surface the midpoint cfvo value so a non-default
                    // percentile (e.g. 40) round-trips instead of silently
                    // resetting to 50. The mid cfvo is the second value object.
                    var csCfvos = colorScale.Elements<ConditionalFormatValueObject>().ToList();
                    if (csCfvos.Count >= 3 && csCfvos[1].Val?.Value is string midValStr
                        && !string.IsNullOrEmpty(midValStr))
                        cfNode.Format["midpoint"] = midValStr;
                }
            }
        }

        // IconSet
        var iconSet = rule.GetFirstChild<IconSet>();
        if (iconSet != null)
        {
            cfNode.Format["type"] = "iconSet";
            if (iconSet.IconSetValue?.Value != null)
                cfNode.Format["iconset"] = iconSet.IconSetValue.InnerText;
            if (iconSet.ShowValue?.Value != null)
                cfNode.Format["showValue"] = iconSet.ShowValue.Value;
            if (iconSet.Reverse?.Value == true)
                cfNode.Format["reverse"] = true;
        }

        // Formula-based
        var formula = rule.GetFirstChild<Formula>();
        if (formula != null && rule.Type?.Value == ConditionalFormatValues.Expression)
        {
            cfNode.Format["type"] = "formula";
            cfNode.Format["formula"] = formula.Text ?? "";
            if (rule.FormatId?.Value != null)
                cfNode.Format["dxfId"] = rule.FormatId.Value;
        }

        // Top/Bottom N
        if (rule.Type?.Value == ConditionalFormatValues.Top10)
        {
            cfNode.Format["type"] = "topN";
            if (rule.Rank?.HasValue == true) cfNode.Format["rank"] = rule.Rank.Value;
            if (rule.Bottom?.Value == true) cfNode.Format["bottom"] = true;
            if (rule.Percent?.Value == true) cfNode.Format["percent"] = true;
            if (rule.FormatId?.Value != null) cfNode.Format["dxfId"] = rule.FormatId.Value;
        }

        // Above/Below Average
        if (rule.Type?.Value == ConditionalFormatValues.AboveAverage)
        {
            cfNode.Format["type"] = "aboveAverage";
            if (rule.AboveAverage?.HasValue == true) cfNode.Format["aboveAverage"] = rule.AboveAverage.Value;
            // stdDev (deviations above/below mean) and equalAverage
            // (include values equal to the mean) round-trip via the
            // cfRule attributes. Only surface when explicitly set.
            if (rule.StdDev?.HasValue == true)
                cfNode.Format["stdDev"] = rule.StdDev.Value;
            if (rule.EqualAverage?.Value == true)
                cfNode.Format["equalAverage"] = true;
            if (rule.FormatId?.Value != null) cfNode.Format["dxfId"] = rule.FormatId.Value;
        }

        // Duplicate Values
        if (rule.Type?.Value == ConditionalFormatValues.DuplicateValues)
        {
            cfNode.Format["type"] = "duplicateValues";
            if (rule.FormatId?.Value != null) cfNode.Format["dxfId"] = rule.FormatId.Value;
        }

        // Unique Values
        if (rule.Type?.Value == ConditionalFormatValues.UniqueValues)
        {
            cfNode.Format["type"] = "uniqueValues";
            if (rule.FormatId?.Value != null) cfNode.Format["dxfId"] = rule.FormatId.Value;
        }

        // Contains Text
        if (rule.Type?.Value == ConditionalFormatValues.ContainsText)
        {
            cfNode.Format["type"] = "containsText";
            if (rule.Text?.HasValue == true) cfNode.Format["text"] = rule.Text.Value;
            if (rule.FormatId?.Value != null) cfNode.Format["dxfId"] = rule.FormatId.Value;
        }

        // Text-operator variants (beginsWith/endsWith/notContainsText) keep
        // their InnerText type from the canonical emit above; surface the
        // rule text so dump can round-trip them like containsText.
        if (rule.Type?.Value == ConditionalFormatValues.BeginsWith
            || rule.Type?.Value == ConditionalFormatValues.EndsWith
            || rule.Type?.Value == ConditionalFormatValues.NotContainsText)
        {
            if (rule.Text?.HasValue == true) cfNode.Format["text"] = rule.Text.Value;
            if (rule.FormatId?.Value != null) cfNode.Format["dxfId"] = rule.FormatId.Value;
        }

        // CellIs (operator-based comparison: between/equal/greaterThan/...)
        if (rule.Type?.Value == ConditionalFormatValues.CellIs)
        {
            cfNode.Format["type"] = "cellIs";
            if (rule.Operator?.HasValue == true)
                cfNode.Format["operator"] = rule.Operator.InnerText;
            var cellIsFormulas = rule.Elements<Formula>().ToList();
            if (cellIsFormulas.Count >= 1)
                cfNode.Format["value"] = cellIsFormulas[0].Text ?? "";
            if (cellIsFormulas.Count >= 2)
                cfNode.Format["value2"] = cellIsFormulas[1].Text ?? "";
            if (rule.FormatId?.Value != null) cfNode.Format["dxfId"] = rule.FormatId.Value;
        }

        // Time Period (date occurring)
        if (rule.Type?.Value == ConditionalFormatValues.TimePeriod)
        {
            cfNode.Format["type"] = "timePeriod";
            if (rule.TimePeriod?.HasValue == true) cfNode.Format["period"] = rule.TimePeriod.InnerText;
            if (rule.FormatId?.Value != null) cfNode.Format["dxfId"] = rule.FormatId.Value;
        }

        // Resolve dxfId to actual fill/font colors from the stylesheet
        if (rule.FormatId?.Value != null)
            PopulateCfNodeFromDxf(cfNode, (int)rule.FormatId.Value);
    }

    /// <summary>
    /// One node per conditional-formatting RULE across all cf elements on a
    /// sheet (a &lt;conditionalFormatting&gt; may hold several &lt;cfRule&gt;
    /// children). The dump emitter iterates this so every rule round-trips;
    /// counting cf ELEMENTS alone dropped the 2nd+ rule of a shared element.
    /// </summary>
    public List<DocumentNode> GetDumpCfRuleNodes(string sheetName)
    {
        var worksheet = FindWorksheet(sheetName)
            ?? throw new System.ArgumentException($"Sheet not found: {sheetName}");
        var result = new List<DocumentNode>();
        foreach (var cf in GetSheet(worksheet).Elements<ConditionalFormatting>())
        {
            var refStr = cf.SequenceOfReferences?.InnerText ?? "";
            foreach (var rule in cf.Elements<ConditionalFormattingRule>())
            {
                var node = new DocumentNode { Type = "conditionalFormatting" };
                node.Format["ref"] = refStr;
                PopulateCfNodeFromRule(worksheet, rule, node);
                result.Add(node);
            }
        }
        return result;
    }
}
