// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Collections.Generic;
using System.Linq;

namespace OfficeCli.Core.Rendering;

/// <summary>
/// Selects an <see cref="IRenderer"/> for a request. Implementations register
/// themselves; the registry never references a specific implementation and asks only
/// "who can serve this?". The dependency points one way — implementations depend on
/// these abstractions, not the reverse — so the default build runs with just the
/// built-in renderer and alternatives can be added independently.
/// <para>
/// Selection rule: among renderers that are <see cref="IRenderer.IsAvailable"/> and
/// whose <see cref="RenderCapabilities.Covers"/> the request, the highest
/// <see cref="RenderCapabilities.Priority"/> wins (registration order breaks ties).
/// </para>
/// </summary>
public sealed class RendererRegistry
{
    /// <summary>Process-wide default registry. The built-in renderer registers here at
    /// startup; additional renderers, when present, register here too.</summary>
    public static RendererRegistry Default { get; } = new();

    private readonly List<IRenderer> _renderers = new();
    private readonly object _gate = new();
    private bool _composed;

    /// <summary>
    /// Raised once per registry, right after the built-in renderers are in place, with the
    /// registry as the argument. Additional renderers add themselves to it here. Subscribe
    /// from a module initializer or composition root and register a singleton instance
    /// (<see cref="Register"/> is idempotent per instance), so no part of the host has to
    /// know the renderer ahead of time.
    /// </summary>
    public static event System.Action<RendererRegistry>? Composing;

    /// <summary>Fire <see cref="Composing"/> for this registry, at most once.</summary>
    internal void RaiseComposingOnce()
    {
        lock (_gate)
        {
            if (_composed) return;
            _composed = true;
        }
        Composing?.Invoke(this);
    }

    /// <summary>Add a renderer. Idempotent per instance (re-adding the same object is a no-op).</summary>
    public void Register(IRenderer renderer)
    {
        if (renderer is null) throw new System.ArgumentNullException(nameof(renderer));
        lock (_gate)
        {
            if (!_renderers.Contains(renderer))
                _renderers.Add(renderer);
        }
    }

    /// <summary>
    /// Best renderer for the request, or null if none can serve it. "Best" = highest
    /// priority among available renderers whose capabilities cover (format, output, mode).
    /// </summary>
    public IRenderer? Resolve(string formatId, RenderOutputKind output, RenderMode mode)
    {
        lock (_gate)
        {
            IRenderer? best = null;
            foreach (var r in _renderers)
            {
                if (!r.IsAvailable) continue;
                if (!r.Capabilities.Covers(formatId, output, mode)) continue;
                if (best is null || r.Capabilities.Priority > best.Capabilities.Priority)
                    best = r;
            }
            return best;
        }
    }

    /// <summary>Snapshot of registered renderers (diagnostics/tests).</summary>
    public IReadOnlyList<RenderCapabilities> Registered()
    {
        lock (_gate) return _renderers.Select(r => r.Capabilities).ToList();
    }
}
