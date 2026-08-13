// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Help;

/// <summary>
/// Fingerprint of the embedded help-schema tree (schemas/help/**), exposed via
/// `officecli crc`. Downstream tooling pins this value and re-checks it after a
/// binary upgrade: unchanged → the documented property surface is identical;
/// changed → re-verify machine-readable output assumptions before resuming
/// automated pipelines. Covers the schema files only — serialization behavior
/// implemented in code (e.g. JSON field order) is outside this fingerprint.
/// </summary>
internal static class SchemaCrc
{
    private static readonly uint[] Table = BuildTable();

    private static uint[] BuildTable()
    {
        var table = new uint[256];
        for (uint i = 0; i < 256; i++)
        {
            uint c = i;
            for (int k = 0; k < 8; k++)
                c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1;
            table[i] = c;
        }
        return table;
    }

    private static uint Append(uint crc, ReadOnlySpan<byte> data)
    {
        foreach (var b in data)
            crc = Table[(crc ^ b) & 0xFF] ^ (crc >> 8);
        return crc;
    }

    /// <summary>
    /// CRC32 over every embedded schemas/help resource, ordered by canonical
    /// name (lowercased, '/'-separated) so the result is stable across MSBuild
    /// path-separator differences. Each entry contributes its canonical name
    /// followed by its raw bytes.
    /// </summary>
    internal static string Compute()
    {
        var asm = typeof(SchemaCrc).Assembly;
        var entries = new List<(string Canonical, string Resource)>();
        foreach (var name in asm.GetManifestResourceNames())
        {
            var canonical = name.Replace('\\', '/').ToLowerInvariant();
            if (canonical.StartsWith("schemas/help/", StringComparison.Ordinal))
                entries.Add((canonical, name));
        }
        entries.Sort((a, b) => string.CompareOrdinal(a.Canonical, b.Canonical));

        uint crc = 0xFFFFFFFFu;
        var buffer = new byte[81920];
        foreach (var (canonical, resource) in entries)
        {
            crc = Append(crc, System.Text.Encoding.UTF8.GetBytes(canonical));
            using var stream = asm.GetManifestResourceStream(resource)!;
            int read;
            while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
                crc = Append(crc, buffer.AsSpan(0, read));
        }
        return (crc ^ 0xFFFFFFFFu).ToString("x8");
    }
}
