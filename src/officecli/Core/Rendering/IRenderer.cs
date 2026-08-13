// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core.Rendering;

/// <summary>
/// What a renderer is handed to render: the already-parsed document, never a file path.
/// The input contract is kept SEPARATE from the renderer so the boundary survives
/// independently of how any one renderer is wired up — a renderer depends only on this
/// interface, never on a concrete handler.
/// </summary>
public interface IRenderInput
{
    /// <summary>Format id of the document, e.g. "docx" / "xlsx" / "pptx".
    /// Used by <see cref="RendererRegistry"/> to match a renderer to the document.</summary>
    string FormatId { get; }

    /// <summary>
    /// The parsed in-memory document, handed over so a renderer works against the
    /// open model rather than re-reading the file. The concrete type is format-specific
    /// (cast by <see cref="FormatId"/>); <c>null</c> when no model is attached.
    /// </summary>
    object? Model { get; }

    /// <summary>Read the node tree at <paramref name="path"/> as format-neutral
    /// <see cref="DocumentNode"/>s. Reflects in-session edits, so a renderer driven off
    /// this stays current as the document is mutated.</summary>
    DocumentNode Get(string path, int depth = 1);

    /// <summary>Query the node tree, format-neutrally.</summary>
    IReadOnlyList<DocumentNode> Query(string selector);
}

/// <summary>
/// A pluggable rendering backend. Implementations are registered with the
/// <see cref="RendererRegistry"/>, which selects one per request by capability and
/// priority; the built-in renderer registers at the default (lowest) priority.
/// <para>
/// Output is artifact-oriented (see <see cref="RenderOutputKind"/>), so an
/// implementation is free to render however it likes — emit HTML for a browser to
/// lay out, emit pre-positioned HTML it laid out itself, or run its own
/// layout+paint engine straight to pixels/PDF.
/// </para>
/// </summary>
public interface IRenderer
{
    /// <summary>Static description used for selection; see <see cref="RenderCapabilities"/>.</summary>
    RenderCapabilities Capabilities { get; }

    /// <summary>Runtime gate. The registry skips renderers that report false,
    /// falling back to the next-highest-priority available renderer.</summary>
    bool IsAvailable { get; }

    /// <summary>Render <paramref name="input"/> per <paramref name="options"/>.</summary>
    RenderResult Render(IRenderInput input, RenderOptions options);
}
