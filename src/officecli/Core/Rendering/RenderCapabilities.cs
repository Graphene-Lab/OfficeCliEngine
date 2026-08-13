// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Collections.Generic;

namespace OfficeCli.Core.Rendering;

/// <summary>
/// What a renderer can do, used by <see cref="RendererRegistry"/> to pick one for a
/// request. The registry chooses the highest-<see cref="Priority"/> available renderer
/// whose capabilities cover the requested format, output kind, and mode.
/// <para>
/// An implementation registers itself with a priority; the registry never references a
/// specific implementation, so alternative renderers can be added independently.
/// </para>
/// </summary>
public sealed record RenderCapabilities
{
    /// <summary>Human-readable id for diagnostics/selection logging (e.g. "basic-html", "native").</summary>
    public required string Name { get; init; }

    /// <summary>Higher wins when several renderers can serve a request. The built-in
    /// basic renderer registers low (0); replacements register above it.</summary>
    public int Priority { get; init; }

    /// <summary>Format ids this renderer handles (e.g. "docx", "xlsx", "pptx").</summary>
    public required IReadOnlyCollection<string> SupportedFormats { get; init; }

    /// <summary>Bitwise-OR of every <see cref="RenderOutputKind"/> this renderer can emit.</summary>
    public required RenderOutputKind SupportedOutputs { get; init; }

    /// <summary>True if the renderer can emit the watch-client interaction contract
    /// (<see cref="RenderMode.Watch"/>): DOM <c>data-path</c> and/or <see cref="RenderResult.HitTest"/>.</summary>
    public bool SupportsWatch { get; init; }

    /// <summary>True if the renderer can populate <see cref="RenderResult.HitTest"/>.</summary>
    public bool SupportsHitTest { get; init; }

    /// <summary>Does this renderer advertise the given format + output + mode?</summary>
    public bool Covers(string formatId, RenderOutputKind output, RenderMode mode)
    {
        if (!SupportedOutputs.HasFlag(output)) return false;
        if (mode == RenderMode.Watch && !SupportsWatch) return false;
        foreach (var f in SupportedFormats)
            if (string.Equals(f, formatId, System.StringComparison.OrdinalIgnoreCase))
                return true;
        return false;
    }
}
