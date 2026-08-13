// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeCli.Core.Diagram;

/// <summary>
/// Layered (Sugiyama) flowchart layout: <see cref="DiagramGraph"/> (semantic IR)
/// → <see cref="LaidOutGraph"/> (geometric IR, coordinates in cm).
///
/// Pipeline (faithful port of the validated Python reference):
///   break cycles (DFS back-edge) → longest-path rank → dummy nodes on long edges
///   → barycenter crossing reduction → Brandes–Köpf cross-axis coordinates
///   → fit-to-canvas (poster sizing) → self-computed orthogonal routing with
///   port distribution and track nudging → self-loops + parallel-label stagger.
///
/// All coordinates emitted in centimetres; the emitter converts cm → EMU.
/// </summary>
public static class FlowchartLayout
{
    private const double VGap = 2.2, HGap = 1.4, EdgeSep = 0.8;

    private sealed class WNode
    {
        public string Id = "";
        public bool Dummy;
        public string Label = "";
        public FlowShape Shape;
        public double X, Y, W, H;
        public int Rank;
    }

    private sealed class WEdge
    {
        public string From = "", To = "", Label = "";
        public bool Rev, Self;
        public List<string> Wp = new();     // waypoint dummy ids, source→target order
        // routing scratch
        public string SSide = "", TSide = "";
        public Pt SRef, TRef;
        public double LabelDy;
    }

    public static LaidOutGraph Layout(DiagramGraph g)
    {
        // A header with no node/edge statements (e.g. a bare "flowchart TD", or
        // input whose only line failed to parse into a node) leaves zero nodes.
        // Guard here with a clear message — otherwise the bounding-box Min/Max
        // below throws a bare "Sequence contains no elements" that surfaces as
        // internal_error. Mirrors SequenceLayout's empty-participants guard.
        if (g.Nodes.Count == 0)
            throw new ArgumentException(
                "diagram has no nodes — the mermaid source has no node/edge statements "
                + "(e.g. 'flowchart TD; A[Start] --> B[End]').");

        bool td = g.Direction == FlowDirection.TopDown;
        var nodes = new Dictionary<string, WNode>();
        var real = new List<string>();
        foreach (var dn in g.Nodes)
        {
            var n = new WNode { Id = dn.Id, Label = dn.Label, Shape = dn.Shape };
            SizeNode(n);
            nodes[dn.Id] = n;
            real.Add(dn.Id);
        }

        var edges = new List<WEdge>();
        foreach (var de in g.Edges)
        {
            if (!nodes.ContainsKey(de.From) || !nodes.ContainsKey(de.To)) continue;
            edges.Add(new WEdge { From = de.From, To = de.To, Label = de.Label, Self = de.From == de.To });
        }

        // 3a. break cycles: DFS, mark back edges (to a node on the stack)
        var adj = Group(edges.Where(e => !e.Self), e => e.From);
        var state = new Dictionary<string, int>();
        void Dfs(string u)
        {
            state[u] = 1;
            if (adj.TryGetValue(u, out var outs))
                foreach (var e in outs)
                {
                    int s = state.GetValueOrDefault(e.To, 0);
                    if (s == 1) e.Rev = true;
                    else if (s == 0) Dfs(e.To);
                }
            state[u] = 2;
        }
        foreach (var i in real)
            if (state.GetValueOrDefault(i, 0) == 0) Dfs(i);

        (string top, string bot) Ends(WEdge e) => e.Rev ? (e.To, e.From) : (e.From, e.To);

        // 3b. longest-path rank (over the DAG, back edges reversed)
        var succ = new Dictionary<string, List<string>>();
        var indeg = real.ToDictionary(i => i, _ => 0);
        foreach (var e in edges.Where(e => !e.Self))
        {
            var (a, b) = Ends(e);
            succ.TryAdd(a, new List<string>()); succ[a].Add(b);
            indeg[b]++;
        }
        var rank = real.ToDictionary(i => i, _ => 0);
        var ind = new Dictionary<string, int>(indeg);
        var q = new Queue<string>(real.Where(i => ind[i] == 0));
        if (q.Count == 0 && real.Count > 0) q.Enqueue(real[0]);
        var seen = new HashSet<string>();
        while (q.Count > 0)
        {
            var u = q.Dequeue();
            if (!seen.Add(u)) continue;
            if (succ.TryGetValue(u, out var vs))
                foreach (var v in vs)
                {
                    rank[v] = Math.Max(rank[v], rank[u] + 1);
                    if (--ind[v] <= 0) q.Enqueue(v);
                }
        }
        int maxRank = rank.Count > 0 ? rank.Values.Max() : 0;
        foreach (var i in real)
            if (!seen.Contains(i)) rank[i] = maxRank;

        // 3c. insert dummy nodes on edges spanning >1 rank
        int dn2 = 0;
        foreach (var e in edges.Where(e => !e.Self))
        {
            var (a, b) = Ends(e);
            var chain = new List<string> { a };
            for (int r = rank[a] + 1; r < rank[b]; r++)
            {
                var did = "__d" + dn2++;
                nodes[did] = new WNode { Id = did, Dummy = true, W = 0.5, H = 0.5, Rank = r };
                chain.Add(did);
            }
            chain.Add(b);
            var wp = chain.GetRange(1, chain.Count - 2);
            if (e.Rev) wp.Reverse();
            e.Wp = wp;
        }
        foreach (var i in real) nodes[i].Rank = rank[i];

        // expanded adjacency (real+dummy) from chains, for barycenter + BK
        var esucc = new Dictionary<string, List<string>>();
        var epred = new Dictionary<string, List<string>>();
        foreach (var e in edges.Where(e => !e.Self))
        {
            var (a, b) = Ends(e);
            var ch = new List<string> { a };
            ch.AddRange(e.Rev ? Enumerable.Reverse(e.Wp) : e.Wp);
            ch.Add(b);
            for (int i = 0; i < ch.Count - 1; i++)
            {
                esucc.TryAdd(ch[i], new List<string>()); esucc[ch[i]].Add(ch[i + 1]);
                epred.TryAdd(ch[i + 1], new List<string>()); epred[ch[i + 1]].Add(ch[i]);
            }
        }
        List<string> Pred(string v) => epred.TryGetValue(v, out var l) ? l : new List<string>();
        List<string> Succ(string v) => esucc.TryGetValue(v, out var l) ? l : new List<string>();

        // ranks → ordered rows
        var allIds = nodes.Keys.ToList();
        var order = new SortedDictionary<int, List<string>>();
        foreach (var i in allIds)
        {
            order.TryAdd(nodes[i].Rank, new List<string>());
            order[nodes[i].Rank].Add(i);
        }

        // barycenter crossing reduction
        var ranks = order.Keys.ToList();
        Dictionary<string, int> PosIn(int r) =>
            order.TryGetValue(r, out var row)
                ? row.Select((n, k) => (n, k)).ToDictionary(t => t.n, t => t.k)
                : new Dictionary<string, int>();
        for (int sweep = 0; sweep < 6; sweep++)
        {
            foreach (var r in ranks.Where(r => r != ranks[0]))
            {
                var p = PosIn(r - 1);
                order[r] = order[r].OrderBy(n => Bary(Pred(n), p)).ToList();
            }
            foreach (var r in Enumerable.Reverse(ranks).Where(r => r != ranks[^1]))
            {
                var p = PosIn(r + 1);
                order[r] = order[r].OrderBy(n => Bary(Succ(n), p)).ToList();
            }
        }

        // LR edge labels sit ALONG the horizontal edge, so a label wider than the
        // inter-rank gap would spill into the adjacent nodes and mask the arrowhead
        // (TD labels are perpendicular to their vertical edge, so VGap already fits
        // them). Reserve extra rank-axis room for the widest label crossing each gap.
        var labelGap = new Dictionary<int, double>();
        if (!td)
            foreach (var e in edges.Where(e => !e.Self && !string.IsNullOrEmpty(e.Label)))
            {
                var (a, b) = Ends(e);
                int gr = Math.Min(rank[a], rank[b]);
                labelGap[gr] = Math.Max(labelGap.GetValueOrDefault(gr, 0), TextExtent(e.Label).w + 1.0);
            }

        // 3d. coordinates: rank-axis = cumulative depth; cross-axis = Brandes–Köpf
        var rankPos = new Dictionary<int, double>();
        double acc = 0;
        foreach (var r in ranks)
        {
            rankPos[r] = acc;
            double span = order[r].Max(j => td ? nodes[j].H : nodes[j].W);
            acc += span + (td ? VGap : HGap) + (td ? 0 : labelGap.GetValueOrDefault(r, 0));
        }
        var layers = ranks.Select(r => order[r]).ToList();
        Func<string, double> csize = td ? (v => nodes[v].W) : (v => nodes[v].H);
        var cross = BkPosition(layers, Pred, Succ, csize, HGap, EdgeSep, id => id.StartsWith("__d"));
        foreach (var r in ranks)
            foreach (var v in order[r])
            {
                double c = cross[v], d = rankPos[r];
                if (td) { nodes[v].X = c - nodes[v].W / 2; nodes[v].Y = d; }
                else { nodes[v].Y = c - nodes[v].H / 2; nodes[v].X = d; }
            }

        // 3e. fit-to-canvas (poster sizing): keep readable; only shrink past PowerPoint max
        double minX = allIds.Min(i => nodes[i].X), minY = allIds.Min(i => nodes[i].Y);
        double maxX = allIds.Max(i => nodes[i].X + nodes[i].W), maxY = allIds.Max(i => nodes[i].Y + nodes[i].H);
        double bw = maxX - minX, bh = maxY - minY;
        const double M = 1.2, MaxD = 55.0;
        double sc = Math.Min(Math.Min(bw > 0 ? MaxD / bw : 1, bh > 0 ? MaxD / bh : 1), 1.0);
        foreach (var i in allIds)
        {
            var n = nodes[i];
            n.X = M + (n.X - minX) * sc; n.Y = M + (n.Y - minY) * sc;
            n.W *= sc; n.H *= sc;
        }

        var outp = new LaidOutGraph
        {
            FontScale = sc,
            SlideWidthCm = Math.Max(bw * sc + 2 * M, 12.0),
            SlideHeightCm = Math.Max(bh * sc + 2 * M, 9.0),
        };
        foreach (var i in real)
        {
            var n = nodes[i];
            outp.Nodes.Add(new PlacedNode { Id = n.Id, Label = n.Label, Shape = n.Shape, X = n.X, Y = n.Y, W = n.W, H = n.H });
        }

        Route(outp, nodes, edges, g, td);
        return outp;
    }

    // ---- routing (produces the geometric IR's polylines + labels) -----------
    private static void Route(LaidOutGraph outp, Dictionary<string, WNode> nodes, List<WEdge> edges,
                              DiagramGraph g, bool td)
    {
        Pt Center(string i) { var n = nodes[i]; return new Pt(n.X + n.W / 2, n.Y + n.H / 2); }
        var valid = edges.Where(e => !e.Self && nodes.ContainsKey(e.From) && nodes.ContainsKey(e.To)).ToList();

        // choose a source/target side by geometry, cache reference coordinate
        foreach (var e in valid)
        {
            var (scx, scy) = (Center(e.From).X, Center(e.From).Y);
            var (tcx, tcy) = (Center(e.To).X, Center(e.To).Y);
            Pt first = e.Wp.Count > 0 ? Center(e.Wp[0]) : new Pt(tcx, tcy);
            Pt last = e.Wp.Count > 0 ? Center(e.Wp[^1]) : new Pt(scx, scy);
            if (td) { e.SSide = first.Y >= scy ? "b" : "t"; e.TSide = last.Y <= tcy ? "t" : "b"; }
            else { e.SSide = first.X >= scx ? "r" : "l"; e.TSide = last.X <= tcx ? "l" : "r"; }
            e.SRef = first; e.TRef = last;
        }

        Pt Attach(WNode node, string side, int k, int tot)
        {
            double frac = (k + 1.0) / (tot + 1.0);
            if (side is "t" or "b")
                return new Pt(node.X + node.W * frac, side == "t" ? node.Y : node.Y + node.H);
            return new Pt(side == "l" ? node.X : node.X + node.W, node.Y + node.H * frac);
        }

        var srcPt = new Dictionary<WEdge, Pt>();
        var tgtPt = new Dictionary<WEdge, Pt>();
        // Distribute ports per (node, side) over BOTH the source ends and target ends
        // that land there. Grouping source- and target-attachments separately (the
        // old bug) gave each a lone-on-its-side centre port, so a node with one
        // incoming AND one outgoing edge on the same side (every cycle: e.g. a state
        // that is both entered and left on its left face) collided both at the centre,
        // producing overlapping collinear segments. One combined group → distinct ports.
        var reqs = new List<(WNode node, string side, WEdge e, bool src, Pt refp)>();
        foreach (var e in valid)
        {
            reqs.Add((nodes[e.From], e.SSide, e, true, e.SRef));
            reqs.Add((nodes[e.To], e.TSide, e, false, e.TRef));
        }
        foreach (var grp in reqs.GroupBy(r => (r.node.Id, r.side)))
        {
            bool horiz = grp.Key.side is "t" or "b"; // horizontal face → order along X
            var rs = grp.OrderBy(r => horiz ? r.refp.X : r.refp.Y).ToList();
            for (int k = 0; k < rs.Count; k++)
            {
                var pt = Attach(rs[k].node, rs[k].side, k, rs.Count);
                if (rs[k].src) srcPt[rs[k].e] = pt; else tgtPt[rs[k].e] = pt;
            }
        }

        // Pass A: build each edge's polyline; jog corners get a nudge-able coordinate
        var routes = new List<List<Pt>>();
        var jogs = new List<Jog>();
        foreach (var e in valid)
        {
            var raw = new List<Pt> { srcPt[e] };
            raw.AddRange(e.Wp.Select(Center));
            raw.Add(tgtPt[e]);
            var pts = new List<Pt> { raw[0] };
            for (int i = 0; i < raw.Count - 1; i++)
            {
                double x1 = pts[^1].X, y1 = pts[^1].Y, x2 = raw[i + 1].X, y2 = raw[i + 1].Y;
                if (Math.Abs(x1 - x2) < 0.05 || Math.Abs(y1 - y2) < 0.05)
                {
                    pts.Add(new Pt(x2, y2));
                }
                else if (td)
                {
                    double m = (y1 + y2) / 2;
                    int ci = pts.Count;
                    pts.Add(new Pt(x1, m)); pts.Add(new Pt(x2, m)); pts.Add(new Pt(x2, y2));
                    jogs.Add(new Jog { Axis = 'y', V = m, A = Math.Min(x1, x2), B = Math.Max(x1, x2), R = routes.Count, C1 = ci, C2 = ci + 1 });
                }
                else
                {
                    double m = (x1 + x2) / 2;
                    int ci = pts.Count;
                    pts.Add(new Pt(m, y1)); pts.Add(new Pt(m, y2)); pts.Add(new Pt(x2, y2));
                    jogs.Add(new Jog { Axis = 'x', V = m, A = Math.Min(y1, y2), B = Math.Max(y1, y2), R = routes.Count, C1 = ci, C2 = ci + 1 });
                }
            }
            routes.Add(pts);
        }

        // Pass B: nudge — split colliding (overlapping-span) jogs in a band onto tracks
        foreach (var band in jogs.GroupBy(j => Math.Round(j.V / 0.8)))
        {
            var js = band.OrderBy(j => j.A).ThenBy(j => j.B).ToList();
            var tracks = new List<List<(double A, double B)>>();
            foreach (var j in js)
            {
                int placed = -1;
                for (int ti = 0; ti < tracks.Count; ti++)
                    if (tracks[ti].All(s => j.B <= s.A || j.A >= s.B)) { tracks[ti].Add((j.A, j.B)); placed = ti; break; }
                if (placed < 0) { tracks.Add(new List<(double, double)> { (j.A, j.B) }); placed = tracks.Count - 1; }
                j.Track = placed;
            }
            int nt = tracks.Count;
            if (nt <= 1) continue;
            double baseV = js.Average(j => j.V);
            foreach (var j in js)
            {
                double newV = baseV + (j.Track - (nt - 1) / 2.0) * 0.6;
                var r = routes[j.R];
                if (j.Axis == 'y') { r[j.C1] = new Pt(r[j.C1].X, newV); r[j.C2] = new Pt(r[j.C2].X, newV); }
                else { r[j.C1] = new Pt(newV, r[j.C1].Y); r[j.C2] = new Pt(newV, r[j.C2].Y); }
            }
        }

        // parallel edges (same from→to) stagger their labels so white masks don't collide
        foreach (var grp in valid.GroupBy(e => (e.From, e.To)))
        {
            var es = grp.ToList();
            for (int k = 0; k < es.Count; k++) es[k].LabelDy = es.Count > 1 ? k * 0.62 : 0.0;
        }

        // Pass C: geometric IR edges + labels
        for (int ei = 0; ei < valid.Count; ei++)
        {
            var e = valid[ei];
            outp.Edges.Add(new RoutedEdge { Points = routes[ei], ArrowAtEnd = true });
            if (!string.IsNullOrEmpty(e.Label))
            {
                var p0 = routes[ei][0]; var p1 = routes[ei][1];
                double dx = p1.X - p0.X, dy = p1.Y - p0.Y;
                double slen = Math.Sqrt(dx * dx + dy * dy); if (slen == 0) slen = 1;
                // Horizontal (LR) segment: centre the label in the reserved gap so its
                // white mask clears both nodes. Otherwise hug the source — branch
                // labels ("Yes"/"No") read best next to the node they leave.
                double t = (Math.Abs(dy) < 0.05 && Math.Abs(dx) > 0.05)
                    ? 0.5
                    : Math.Min(0.85, slen * 0.45) / slen;
                outp.Labels.Add(new EdgeLabel { Text = e.Label, Cx = p0.X + dx * t, Cy = p0.Y + dy * t + e.LabelDy });
            }
        }

        // self-loops: a small side loop with the arrowhead back into the node
        foreach (var e in edges.Where(e => e.Self && nodes.ContainsKey(e.From)))
        {
            var n = nodes[e.From]; const double L = 1.0;
            var pts = new List<Pt>();
            Pt lp;
            if (td)
            {
                double rx = n.X + n.W, a = n.Y + n.H * 0.3, b = n.Y + n.H * 0.7;
                pts.Add(new Pt(rx, a)); pts.Add(new Pt(rx + L, a)); pts.Add(new Pt(rx + L, b)); pts.Add(new Pt(rx, b));
                lp = new Pt(rx + L + 0.7, (a + b) / 2);
            }
            else
            {
                double by = n.Y + n.H, a = n.X + n.W * 0.3, b = n.X + n.W * 0.7;
                pts.Add(new Pt(a, by)); pts.Add(new Pt(a, by + L)); pts.Add(new Pt(b, by + L)); pts.Add(new Pt(b, by));
                lp = new Pt((a + b) / 2, by + L + 0.3);
            }
            outp.Edges.Add(new RoutedEdge { Points = pts, ArrowAtEnd = true });
            if (!string.IsNullOrEmpty(e.Label))
                outp.Labels.Add(new EdgeLabel { Text = e.Label, Cx = lp.X, Cy = lp.Y });
        }
    }

    private sealed class Jog
    {
        public char Axis;
        public double V, A, B;
        public int R, C1, C2, Track;
    }

    // ---- Brandes–Köpf cross-axis coordinate assignment (dagre port) ---------
    private static Dictionary<string, double> BkPosition(
        List<List<string>> layers, Func<string, List<string>> pred, Func<string, List<string>> succ,
        Func<string, double> csize, double nodesep, double edgesep, Func<string, bool> isDummy)
    {
        var order = new Dictionary<string, int>();
        foreach (var lyr in layers)
            for (int i = 0; i < lyr.Count; i++) order[lyr[i]] = i;

        var conflicts = new HashSet<(string, string)>();
        void AddC(string v, string w) => conflicts.Add(string.CompareOrdinal(v, w) < 0 ? (v, w) : (w, v));
        bool HasC(string v, string w) => conflicts.Contains(string.CompareOrdinal(v, w) < 0 ? (v, w) : (w, v));
        string? Inner(string v)
        {
            if (isDummy(v))
                foreach (var u in pred(v))
                    if (isDummy(u)) return u;
            return null;
        }
        for (int li = 1; li < layers.Count; li++)
        {
            var prev = layers[li - 1]; var lyr = layers[li];
            int k0 = 0, scanPos = 0, plen = prev.Count;
            string? last = lyr.Count > 0 ? lyr[^1] : null;
            for (int i = 0; i < lyr.Count; i++)
            {
                var v = lyr[i];
                var w = Inner(v);
                int k1 = w != null ? order[w] : plen;
                if (w != null || v == last)
                {
                    for (int s = scanPos; s <= i; s++)
                    {
                        var sn = lyr[s];
                        foreach (var u in pred(sn))
                        {
                            int up = order[u];
                            if ((up < k0 || k1 < up) && !(isDummy(u) && isDummy(sn))) AddC(u, sn);
                        }
                    }
                    scanPos = i + 1; k0 = k1;
                }
            }
        }

        (Dictionary<string, string> root, Dictionary<string, string> align) VAlign(
            List<List<string>> lays, Func<string, List<string>> neigh)
        {
            var root = new Dictionary<string, string>();
            var align = new Dictionary<string, string>();
            var pos = new Dictionary<string, int>();
            foreach (var lyr in lays)
                for (int o = 0; o < lyr.Count; o++) { root[lyr[o]] = lyr[o]; align[lyr[o]] = lyr[o]; pos[lyr[o]] = o; }
            foreach (var lyr in lays)
            {
                int prevIdx = -1;
                foreach (var v in lyr)
                {
                    var ws = neigh(v).OrderBy(a => pos.GetValueOrDefault(a, 0)).ToList();
                    if (ws.Count == 0) continue;
                    double mp = (ws.Count - 1) / 2.0;
                    for (int i = (int)Math.Floor(mp); i <= (int)Math.Ceiling(mp); i++)
                    {
                        var w = ws[i];
                        if (align[v] == v && prevIdx < pos[w] && !HasC(v, w))
                        {
                            align[w] = v; align[v] = root[v] = root[w]; prevIdx = pos[w];
                        }
                    }
                }
            }
            return (root, align);
        }

        double Sep(string v, string u) =>
            (csize(v) + csize(u)) / 2
            + (isDummy(v) ? edgesep : nodesep) / 2 + (isDummy(u) ? edgesep : nodesep) / 2;

        Dictionary<string, double> Compact(List<List<string>> lays, Dictionary<string, string> root, Dictionary<string, string> align)
        {
            var bnodes = new HashSet<string>();
            var bedge = new Dictionary<(string, string), double>();
            foreach (var lyr in lays)
            {
                string? u = null;
                foreach (var v in lyr)
                {
                    var rv = root[v]; bnodes.Add(rv);
                    if (u != null)
                    {
                        var ru = root[u]; var key = (ru, rv);
                        bedge[key] = Math.Max(Sep(v, u), bedge.GetValueOrDefault(key, 0));
                    }
                    u = v;
                }
            }
            var bin = new Dictionary<string, List<(string, double)>>();
            var bout = new Dictionary<string, List<(string, double)>>();
            foreach (var kv in bedge)
            {
                var (a, b) = kv.Key;
                bout.TryAdd(a, new()); bout[a].Add((b, kv.Value));
                bin.TryAdd(b, new()); bin[b].Add((a, kv.Value));
            }
            var xs = new Dictionary<string, double>();
            void Iterate(Action<string> setx, Func<string, List<(string, double)>> next)
            {
                var stack = new Stack<string>(bnodes);
                var vis = new HashSet<string>();
                while (stack.Count > 0)
                {
                    var e = stack.Pop();
                    if (vis.Contains(e)) setx(e);
                    else { vis.Add(e); stack.Push(e); foreach (var (nx, _) in next(e)) stack.Push(nx); }
                }
            }
            List<(string, double)> In(string e) => bin.TryGetValue(e, out var l) ? l : new();
            List<(string, double)> Out(string e) => bout.TryGetValue(e, out var l) ? l : new();
            Iterate(e => xs[e] = In(e).Count == 0 ? 0 : In(e).Max(t => xs.GetValueOrDefault(t.Item1, 0) + t.Item2), In);
            Iterate(e =>
            {
                var outs = Out(e);
                if (outs.Count > 0) xs[e] = Math.Max(xs.GetValueOrDefault(e, 0), outs.Min(t => xs.GetValueOrDefault(t.Item1, 0) - t.Item2));
            }, Out);
            var res = new Dictionary<string, double>();
            foreach (var v in align.Keys) res[v] = xs.GetValueOrDefault(root[v], 0);
            return res;
        }

        var xss = new Dictionary<string, Dictionary<string, double>>();
        foreach (var vert in new[] { "u", "d" })
        {
            var baseL = vert == "u" ? layers : Enumerable.Reverse(layers).ToList();
            foreach (var horiz in new[] { "l", "r" })
            {
                var lays = horiz == "r"
                    ? baseL.Select(l => Enumerable.Reverse(l).ToList()).ToList()
                    : baseL.Select(l => l.ToList()).ToList();
                Func<string, List<string>> neigh = vert == "u" ? pred : succ;
                var (root, align) = VAlign(lays, neigh);
                var xs = Compact(lays, root, align);
                if (horiz == "r") xs = xs.ToDictionary(k => k.Key, k => -k.Value);
                xss[vert + horiz] = xs;
            }
        }

        double Width(Dictionary<string, double> xs)
        {
            double mx = double.NegativeInfinity, mn = double.PositiveInfinity;
            foreach (var (v, x) in xs) { mx = Math.Max(mx, x + csize(v) / 2); mn = Math.Min(mn, x - csize(v) / 2); }
            return mx - mn;
        }
        var small = xss.Values.OrderBy(Width).First();
        double smin = small.Values.Min(), smax = small.Values.Max();
        foreach (var key in xss.Keys.ToList())
        {
            var xs = xss[key];
            if (ReferenceEquals(xs, small)) continue;
            double delta = key[1] == 'l' ? smin - xs.Values.Min() : smax - xs.Values.Max();
            if (delta != 0) xss[key] = xs.ToDictionary(k => k.Key, k => k.Value + delta);
        }
        var outd = new Dictionary<string, double>();
        foreach (var v in xss["ul"].Keys)
        {
            var vals = new[] { xss["ul"][v], xss["ur"][v], xss["dl"][v], xss["dr"][v] };
            Array.Sort(vals);
            outd[v] = (vals[1] + vals[2]) / 2;
        }
        return outd;
    }

    // ---- sizing -------------------------------------------------------------
    private static (double w, int lines) TextExtent(string label)
    {
        double w = 0;
        foreach (var c in label) w += c > 0x2E80 ? 0.58 : 0.30;
        const double maxLine = 5.0;
        int lines = Math.Max(1, (int)(w / maxLine) + (w % maxLine != 0 ? 1 : 0));
        return (Math.Min(w, maxLine), lines);
    }

    private static void SizeNode(WNode n)
    {
        var (tw, lines) = TextExtent(n.Label);
        double w = tw + 1.0, h = 0.7 + lines * 0.62;
        switch (n.Shape)
        {
            case FlowShape.Decision: w = tw * 2.2 + 1.0; h = Math.Max(lines * 1.24 + 0.9, 2.2); break;
            case FlowShape.Hexagon: w = tw + 1.8; h = Math.Max(lines * 0.62 + 0.9, 1.4); break;
            case FlowShape.Parallelogram: w = tw + 1.8; break;
            case FlowShape.Database: h += 0.7; break;
            case FlowShape.Circle: { double s = Math.Max(w, h) * 1.35; w = s; h = s; break; }
        }
        n.W = Math.Max(w, 2.4); n.H = Math.Max(h, 1.1);
    }

    private static double Bary(List<string> ns, Dictionary<string, int> pos) =>
        ns.Count == 0 ? 1e9 : ns.Average(x => (double)pos.GetValueOrDefault(x, 0));

    private static Dictionary<string, List<WEdge>> Group(IEnumerable<WEdge> es, Func<WEdge, string> key)
    {
        var d = new Dictionary<string, List<WEdge>>();
        foreach (var e in es) { var k = key(e); d.TryAdd(k, new()); d[k].Add(e); }
        return d;
    }
}
