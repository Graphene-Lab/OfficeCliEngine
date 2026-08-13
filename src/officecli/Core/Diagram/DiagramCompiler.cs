// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System;
using System.Text.RegularExpressions;

namespace OfficeCli.Core.Diagram;

/// <summary>
/// Single entry point for `add --type diagram`: sniffs the mermaid header and
/// dispatches to the matching layout engine, returning the shared geometric IR
/// (<see cref="LaidOutGraph"/>). Mermaid diagram types we don't render yet are
/// rejected with a clear message rather than producing garbage — the `diagram`
/// umbrella name stays honest.
/// </summary>
public static class DiagramCompiler
{
    public static LaidOutGraph Compile(string mermaid)
    {
        var header = FirstMeaningfulLine(mermaid);

        if (Regex.IsMatch(header, @"^(flowchart|graph)\b", RegexOptions.IgnoreCase))
            return FlowchartLayout.Layout(MermaidParser.Parse(mermaid));

        if (Regex.IsMatch(header, @"^sequenceDiagram\b", RegexOptions.IgnoreCase))
            return SequenceLayout.Layout(SequenceLayout.Parse(mermaid));

        // No explicit header → assume flowchart (mermaid's own lenient default).
        if (header.Length == 0 || !Regex.IsMatch(header, @"^[A-Za-z]"))
            return FlowchartLayout.Layout(MermaidParser.Parse(mermaid));

        var kind = Regex.Match(header, @"^[A-Za-z]+").Value;
        throw new ArgumentException(
            $"diagram type '{kind}' is not supported yet (currently: flowchart, sequenceDiagram).");
    }

    private static string FirstMeaningfulLine(string text)
    {
        foreach (var raw in text.Split('\n'))
        {
            var s = raw.Trim();
            if (s.Length > 0 && !s.StartsWith("%%"))
                return s;
        }
        return "";
    }
}
