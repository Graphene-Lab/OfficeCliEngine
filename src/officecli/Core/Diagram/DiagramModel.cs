// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Collections.Generic;

namespace OfficeCli.Core.Diagram;

/// <summary>Flowchart node shape kind (mermaid shape token → drawing preset).</summary>
public enum FlowShape
{
    Process,       // [text]      rectangle
    Decision,      // {text}      diamond
    Terminator,    // (text)      rounded rectangle
    Stadium,       // ([text])    pill
    Circle,        // ((text))    ellipse
    Hexagon,       // {{text}}    hexagon
    Parallelogram, // [/text/]    parallelogram
    Database,      // [(text)]    cylinder
    Subroutine,    // [[text]]    framed rectangle
    Flag,          // >text]      asymmetric
}

public enum FlowDirection { TopDown, LeftRight }

// ---- Semantic IR (front-end boundary: logical graph, NO coordinates) --------
// This is what mermaid / draw.io / graphviz-dot all map onto. The layout engine
// consumes it; front-ends produce it. Serializable so `add diagram` can also
// accept an IR object instead of mermaid text.

public sealed class DiagramNode
{
    public string Id = "";
    public string Label = "";
    public FlowShape Shape = FlowShape.Process;
}

public sealed class DiagramEdge
{
    public string From = "";
    public string To = "";
    public string Label = "";
}

public sealed class DiagramGraph
{
    /// <summary>Nodes in first-seen order (matched by the emitter's positional index).</summary>
    public readonly List<DiagramNode> Nodes = new();
    public readonly Dictionary<string, DiagramNode> NodeById = new();
    public readonly List<DiagramEdge> Edges = new();
    public FlowDirection Direction = FlowDirection.TopDown;

    public DiagramNode GetOrAdd(string id)
    {
        if (!NodeById.TryGetValue(id, out var n))
        {
            n = new DiagramNode { Id = id, Label = id };
            NodeById[id] = n;
            Nodes.Add(n);
        }
        return n;
    }
}

// ---- Geometric IR (back-end boundary: laid-out, coordinates in cm) ----------
// The layout engine produces it; any emitter (pptx / docx / svg) consumes it.
// draw.io input (which already carries coordinates) would enter at THIS layer,
// skipping the layout engine — that is the whole point of the two-IR split.

public readonly struct Pt
{
    public readonly double X;
    public readonly double Y;
    public Pt(double x, double y) { X = x; Y = y; }
}

public sealed class PlacedNode
{
    public string Id = "";
    public string Label = "";
    public FlowShape Shape;
    public double X, Y, W, H; // cm, top-left + size
}

/// <summary>An edge as an orthogonal polyline; the emitter draws one straight
/// connector per segment and puts the arrowhead on the final segment.</summary>
public sealed class RoutedEdge
{
    public List<Pt> Points = new();
    public bool ArrowAtEnd = true;
    public bool Dashed;          // dashed stroke (sequence lifelines & return messages)
}

public sealed class EdgeLabel
{
    public string Text = "";
    public double Cx, Cy; // cm, center
    // Flowchart labels sit ON the edge line and need an opaque (white) backing
    // to mask it. Sequence-message labels sit ABOVE the arrow in empty space, so
    // an opaque backing would only mask whatever lifeline it overlaps → set false.
    public bool Opaque = true;
}

public sealed class LaidOutGraph
{
    public readonly List<PlacedNode> Nodes = new();
    public readonly List<RoutedEdge> Edges = new();
    public readonly List<EdgeLabel> Labels = new();
    public double SlideWidthCm;
    public double SlideHeightCm;
    public double FontScale = 1.0;
}
