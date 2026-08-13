// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Collections.Generic;

namespace OfficeCli.Core.Rendering;

/// <summary>
/// One rendered region mapped back to a document path. This is the interaction
/// contract in its renderer-neutral form: HTML renderers can express the same thing
/// via <c>data-path</c> in the DOM (and may leave <see cref="RenderResult.HitTest"/>
/// null); native/canvas renderers, which have no DOM, emit these regions so a client
/// can hit-test selection/marks against the painted surface (as a canvas-based
/// editor does).
/// </summary>
public sealed record HitTestRegion(
    int Page,
    double X,
    double Y,
    double Width,
    double Height,
    string DocumentPath);

/// <summary>
/// The product of a render. Carries exactly one artifact (text for Html/Svg, bytes for
/// Png/Pdf) plus optional metadata. Use the static factories rather than the constructor.
/// </summary>
public sealed class RenderResult
{
    /// <summary>The artifact kind actually produced.</summary>
    public required RenderOutputKind Output { get; init; }

    /// <summary>Text artifact for <see cref="RenderOutputKind.Html"/> / <see cref="RenderOutputKind.Svg"/>.</summary>
    public string? Text { get; init; }

    /// <summary>Binary artifact for <see cref="RenderOutputKind.Png"/> / <see cref="RenderOutputKind.Pdf"/>.</summary>
    public byte[]? Bytes { get; init; }

    /// <summary>Total page count of the source document, when the renderer computes it.</summary>
    public int? PageCount { get; init; }

    /// <summary>Optional region -> document-path map (see <see cref="HitTestRegion"/>).
    /// Null is valid: HTML renderers usually encode interaction in the DOM instead.</summary>
    public IReadOnlyList<HitTestRegion>? HitTest { get; init; }

    public static RenderResult Html(string html, int? pageCount = null, IReadOnlyList<HitTestRegion>? hitTest = null)
        => new() { Output = RenderOutputKind.Html, Text = html, PageCount = pageCount, HitTest = hitTest };

    public static RenderResult Svg(string svg, int? pageCount = null, IReadOnlyList<HitTestRegion>? hitTest = null)
        => new() { Output = RenderOutputKind.Svg, Text = svg, PageCount = pageCount, HitTest = hitTest };

    public static RenderResult Png(byte[] bytes, int? pageCount = null, IReadOnlyList<HitTestRegion>? hitTest = null)
        => new() { Output = RenderOutputKind.Png, Bytes = bytes, PageCount = pageCount, HitTest = hitTest };

    public static RenderResult Pdf(byte[] bytes, int? pageCount = null)
        => new() { Output = RenderOutputKind.Pdf, Bytes = bytes, PageCount = pageCount };
}
