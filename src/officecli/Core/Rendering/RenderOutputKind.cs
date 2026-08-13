// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core.Rendering;

/// <summary>
/// The artifact a renderer produces. Declared as [Flags] so a renderer's
/// <see cref="RenderCapabilities.SupportedOutputs"/> can advertise several at once,
/// while a single <see cref="RenderOptions.Output"/> request names exactly one.
/// <para>
/// The contract is deliberately artifact-oriented (Html / Svg / Png / Pdf), NOT
/// mechanism-oriented. An HTML-based renderer and a self-contained
/// layout/paint renderer both satisfy the same interface; HTML is just one possible
/// output, not a requirement.
/// </para>
/// </summary>
[System.Flags]
public enum RenderOutputKind
{
    None = 0,
    /// <summary>HTML markup (text). The basic renderer's native output; carries
    /// watch-client annotations when <see cref="RenderMode.Watch"/> is requested.</summary>
    Html = 1,
    /// <summary>Standalone SVG markup (text).</summary>
    Svg = 2,
    /// <summary>Rasterized PNG (bytes). HTML-based renderers reach this by piping HTML
    /// through a headless browser; native renderers paint it directly.</summary>
    Png = 4,
    /// <summary>PDF (bytes).</summary>
    Pdf = 8,
}

/// <summary>
/// Whether the rendered output must carry the watch-client interaction contract
/// (stable element/region identity -> document path, scroll anchors, math hooks).
/// </summary>
public enum RenderMode
{
    /// <summary>One-shot output for view/screenshot/export. No interaction annotations.</summary>
    Static = 0,
    /// <summary>Live-preview output for the watch loop. The result MUST expose the
    /// interaction contract: HTML renderers via <c>data-path</c> attributes, native
    /// renderers via <see cref="RenderResult.HitTest"/>. See the watch-contract note
    /// in this folder's README.</summary>
    Watch = 1,
}
