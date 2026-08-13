// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Collections.Generic;

namespace OfficeCli.Core.Diagram;

/// <summary>
/// Format-agnostic visual style for a laid-out diagram: the OOXML preset
/// geometry name + fill/line colors per <see cref="FlowShape"/>, plus the edge
/// color. Shared by the pptx and docx emitters so the two never drift (the
/// geometry strings — "rect", "diamond", "can", … — are valid DrawingML
/// <c>a:prstGeom@prst</c> values usable directly in docx and via
/// <c>TryParsePresetShape</c> in pptx).
/// </summary>
public static class DiagramStyles
{
    public static readonly IReadOnlyDictionary<FlowShape, (string Geometry, string Fill, string Line)> ByShape =
        new Dictionary<FlowShape, (string, string, string)>
        {
            [FlowShape.Process]       = ("rect",          "DAE8FC", "6C8EBF"),
            [FlowShape.Decision]      = ("diamond",       "FFF2CC", "D6B656"),
            [FlowShape.Terminator]    = ("roundRect",     "D5E8D4", "82B366"),
            [FlowShape.Stadium]       = ("roundRect",     "D5E8D4", "82B366"),
            [FlowShape.Circle]        = ("ellipse",       "F8CECC", "B85450"),
            [FlowShape.Hexagon]       = ("hexagon",       "FFF2CC", "D6B656"),
            [FlowShape.Parallelogram] = ("parallelogram", "DAE8FC", "6C8EBF"),
            [FlowShape.Database]      = ("can",           "E1D5E7", "9673A6"),
            [FlowShape.Subroutine]    = ("rect",          "DAE8FC", "6C8EBF"),
            [FlowShape.Flag]          = ("rect",          "DAE8FC", "6C8EBF"),
        };

    /// <summary>Connector / edge stroke color (dark grey).</summary>
    public const string EdgeColor = "4D4D4D";
}
