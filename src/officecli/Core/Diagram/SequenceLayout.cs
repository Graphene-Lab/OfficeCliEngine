// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace OfficeCli.Core.Diagram;

// ---- semantic IR (sequence flavor) ------------------------------------------
public sealed class SeqParticipant { public string Id = ""; public string Label = ""; }

public sealed class SeqMessage
{
    public string From = "", To = "", Label = "";
    public bool Dashed;   // dotted line (mermaid `--`) — conventionally a return
    public bool Arrow;    // has an arrowhead (`>>`, `>`, `x`, `)`)
}

public sealed class SequenceDiagram
{
    public readonly List<SeqParticipant> Participants = new();
    public readonly Dictionary<string, SeqParticipant> ById = new();
    public readonly List<SeqMessage> Messages = new();

    public SeqParticipant See(string id)
    {
        if (!ById.TryGetValue(id, out var p))
        {
            p = new SeqParticipant { Id = id, Label = id };
            ById[id] = p; Participants.Add(p);
        }
        return p;
    }
}

/// <summary>
/// Mermaid sequenceDiagram subset → lifeline layout → shared <see cref="LaidOutGraph"/>.
/// Participants become box nodes with a dashed lifeline; messages are horizontal
/// arrows stacked top→bottom (solid call / dashed return); self-messages loop.
/// Deferred: activation bars, alt/opt/loop fragments, notes.
/// </summary>
public static class SequenceLayout
{
    // Participant/actor ids accept Unicode letters, not just ASCII — a
    // fully-Chinese sequence (客户->>服务器: 登录) must parse, mirroring the
    // flowchart parser. \p{L} = any letter, \p{N} = any digit.
    private const string SeqId = @"[\p{L}\p{N}_]+";
    private static readonly Regex Decl =
        new(@"^(?:participant|actor)\s+(" + SeqId + @")(?:\s+as\s+(.+))?$", RegexOptions.IgnoreCase);
    // `A->>B: msg`, plus optional activation control `+`/`-` on the target
    // (`A->>+B`, `B-->>-A`) — we don't draw activation bars, but the message must
    // still render rather than being dropped.
    private static readonly Regex Msg =
        new(@"^(" + SeqId + @")\s*(-{1,2}[>)x]{1,2})\s*[+-]?\s*(" + SeqId + @")\s*:\s*(.*)$");

    public static SequenceDiagram Parse(string text)
    {
        var d = new SequenceDiagram();
        // Split on ';' as well as newlines so the single-line form
        // ("sequenceDiagram; A->>B: hi; B-->>A: ok") parses — same statement
        // separator the flowchart parser already accepts.
        foreach (var raw in text.Split('\n', ';'))
        {
            var line = raw.Trim();
            if (line.Length == 0 || line.StartsWith("%%") ||
                Regex.IsMatch(line, @"^sequenceDiagram\b", RegexOptions.IgnoreCase))
                continue;

            var md = Decl.Match(line);
            if (md.Success)
            {
                var p = d.See(md.Groups[1].Value);
                if (md.Groups[2].Success) p.Label = md.Groups[2].Value.Trim();
                continue;
            }
            var mm = Msg.Match(line);
            if (mm.Success)
            {
                var op = mm.Groups[2].Value;
                d.See(mm.Groups[1].Value); d.See(mm.Groups[3].Value);
                d.Messages.Add(new SeqMessage
                {
                    From = mm.Groups[1].Value,
                    To = mm.Groups[3].Value,
                    Label = mm.Groups[4].Value.Trim(),
                    Dashed = op.StartsWith("--"),
                    Arrow = op.Contains('>') || op.Contains('x') || op.Contains(')'),
                });
            }
        }
        return d;
    }

    public static LaidOutGraph Layout(SequenceDiagram d)
    {
        const double boxH = 1.1, top = 0.8, hGap = 1.4, row = 1.15;
        var order = d.Participants;
        var lo = new LaidOutGraph { FontScale = 1.0 };
        if (order.Count == 0)
            throw new ArgumentException("sequence diagram has no participants.");

        // participant x positions (left, width) + lifeline centre
        var left = new Dictionary<string, double>();
        var width = new Dictionary<string, double>();
        var cxOf = new Dictionary<string, double>();
        double cur = 0.8;
        foreach (var p in order)
        {
            double w = Math.Max(2.4, TextWidth(p.Label) + 1.0);
            left[p.Id] = cur; width[p.Id] = w; cxOf[p.Id] = cur + w / 2;
            cur += w + hGap;
        }
        double bodyTop = top + boxH + 0.9;
        double bottom = bodyTop + Math.Max(1, d.Messages.Count) * row + 0.6;
        lo.SlideWidthCm = Math.Max(cur - hGap + 0.8, 12.0);
        lo.SlideHeightCm = bottom + 0.8;

        // participant boxes + lifelines
        foreach (var p in order)
        {
            lo.Nodes.Add(new PlacedNode { Id = p.Id, Label = p.Label, Shape = FlowShape.Process,
                X = left[p.Id], Y = top, W = width[p.Id], H = boxH });
            lo.Edges.Add(new RoutedEdge
            {
                Dashed = true, ArrowAtEnd = false,
                Points = new List<Pt> { new(cxOf[p.Id], top + boxH), new(cxOf[p.Id], bottom) },
            });
        }

        // messages
        for (int i = 0; i < d.Messages.Count; i++)
        {
            var m = d.Messages[i];
            double y = bodyTop + i * row;
            double x1 = cxOf[m.From], x2 = cxOf[m.To];
            if (m.From == m.To)
            {
                double r = x1 + 1.4;
                lo.Edges.Add(new RoutedEdge { ArrowAtEnd = true, Points = new List<Pt>
                    { new(x1, y), new(r, y), new(r, y + 0.45), new(x1, y + 0.45) } });
                if (m.Label.Length > 0)
                    lo.Labels.Add(new EdgeLabel { Text = m.Label, Cx = x1 + 1.0, Cy = y - 0.25, Opaque = false });
            }
            else
            {
                lo.Edges.Add(new RoutedEdge { ArrowAtEnd = m.Arrow, Dashed = m.Dashed,
                    Points = new List<Pt> { new(x1, y), new(x2, y) } });
                if (m.Label.Length > 0)
                    lo.Labels.Add(new EdgeLabel { Text = m.Label, Cx = (x1 + x2) / 2, Cy = y - 0.5, Opaque = false });
            }
        }
        return lo;
    }

    private static double TextWidth(string s)
    {
        double w = 0;
        foreach (var c in s) w += c > 0x2E80 ? 0.58 : 0.30;
        return w;
    }
}
