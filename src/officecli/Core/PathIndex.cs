// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// The single, named boundary between the two positional bases in the codebase:
///
///   • path / xpath-index — 1-based. The <c>[N]</c> positions in a path segment
///     (<c>/slide[1]/shape[2]/paragraph[1]/run[1]</c>), XPath-style, and the
///     <c>slideIdx</c>/<c>shapeIdx</c>/<c>paraIdx</c>/… locals parsed from them.
///   • array index — 0-based. Any list/array position.
///
/// Wrap the conversion so a reader sees the boundary explicitly instead of a bare
/// <c>- 1</c> / <c>+ 1</c> whose base is implicit. Use it ONLY for a genuine
/// base conversion (indexing a list by a 1-based path position, or naming an
/// array position as a 1-based path position) — NOT for 1-based arithmetic that
/// stays in the path world (e.g. an OOXML row shift <c>rowIdx - 1</c> that means
/// "the row above", still a 1-based row number).
/// </summary>
internal static class PathIndex
{
    /// <summary>1-based xpath position → 0-based array index (<c>n - 1</c>).</summary>
    public static int ToArrayIndex(int xpathIndex) => xpathIndex - 1;

    /// <summary>0-based array index → 1-based xpath position (<c>n + 1</c>).</summary>
    public static int FromArrayIndex(int arrayIndex) => arrayIndex + 1;
}
