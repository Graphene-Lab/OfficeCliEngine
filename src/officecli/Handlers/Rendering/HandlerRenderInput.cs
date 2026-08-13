// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Collections.Generic;
using OfficeCli.Core;
using OfficeCli.Core.Rendering;

namespace OfficeCli.Handlers.Rendering;

/// <summary>
/// A handler that can hand out its parsed in-memory model for rendering. Implemented by
/// the native handlers; lets a render input expose the model without widening the
/// mutation-facing <see cref="IDocumentHandler"/> surface.
/// </summary>
internal interface IRenderModelHost
{
    object? RenderModel { get; }
}

/// <summary>
/// Binds a live handler to the <see cref="IRenderInput"/> seam. Read access (the parsed
/// model and the node tree) is projected off the handler; the renderer never sees the
/// mutation surface or the file path.
/// </summary>
public sealed class HandlerRenderInput : IRenderInput
{
    public IDocumentHandler Handler { get; }
    public string FormatId { get; }

    public HandlerRenderInput(IDocumentHandler handler, string formatId)
    {
        Handler = handler;
        FormatId = formatId;
    }

    public object? Model => (Handler as IRenderModelHost)?.RenderModel;

    public DocumentNode Get(string path, int depth = 1) => Handler.Get(path, depth);

    public IReadOnlyList<DocumentNode> Query(string selector) => Handler.Query(selector);
}
