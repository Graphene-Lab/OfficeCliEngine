// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Net;
using System.Net.Http;
using System.Net.Sockets;

namespace OfficeCli.Core;

/// <summary>
/// Shared SSRF protection for every remote (HTTP/HTTPS) fetch in the tool.
///
/// Both <see cref="ImageSource"/> (picture=) and <see cref="FileSource"/>
/// (table data=, model3d=, media=) pull bytes from caller-supplied URLs. When
/// officecli runs as an agent/automation tool, that URL can originate from
/// untrusted input (a batch script, an instruction embedded in a document,
/// a tool-call argument), so an unguarded fetch is an SSRF primitive: probe
/// intranet hosts / cloud-metadata endpoints, or exfiltrate their bodies into
/// the produced document. This guard makes every remote fetch refuse to land
/// on a non-public address. Keep image and file fetch on this one helper so the
/// policy can never drift apart between them again.
/// </summary>
internal static class SsrfGuard
{
    /// <summary>
    /// Build an <see cref="HttpMessageHandler"/> that validates the *actual* IP
    /// at connect time, for every connection including each redirect hop.
    /// Redirects stay enabled so legitimate public CDNs that 30x still work, but
    /// no hop is allowed to land on a loopback / private / link-local address
    /// (cloud metadata, localhost services, intranet hosts). Validating in
    /// ConnectCallback — rather than resolving the hostname up front — also
    /// closes the DNS-rebinding/TOCTOU window, since the address we vet is the
    /// address we connect to.
    /// </summary>
    /// <param name="what">Noun used in the refusal message, e.g. "image" or "file".</param>
    public static SocketsHttpHandler CreateGuardedHandler(string what)
    {
        return new SocketsHttpHandler
        {
            AllowAutoRedirect = true,
            MaxAutomaticRedirections = 10,
            ConnectCallback = async (ctx, ct) =>
            {
                var host = ctx.DnsEndPoint.Host;
                var addresses = await Dns.GetHostAddressesAsync(host, ct).ConfigureAwait(false);
                foreach (var addr in addresses)
                {
                    if (!IsPublicAddress(addr))
                        throw new ArgumentException(
                            $"Refusing to fetch {what} from non-public address '{addr}' (host '{host}'). " +
                            $"Remote {what} sources must resolve to a public IP (SSRF protection).");
                }
                var target = addresses.FirstOrDefault()
                    ?? throw new ArgumentException($"Could not resolve host '{host}'.");
                var socket = new Socket(SocketType.Stream, ProtocolType.Tcp) { NoDelay = true };
                try
                {
                    await socket.ConnectAsync(target, ctx.DnsEndPoint.Port, ct).ConfigureAwait(false);
                    return new NetworkStream(socket, ownsSocket: true);
                }
                catch
                {
                    socket.Dispose();
                    throw;
                }
            }
        };
    }

    /// <summary>
    /// Shared cap on a single remote fetch, to bound memory use on a hostile or
    /// accidentally-huge response. Without it the default HttpClient buffering
    /// ceiling (~2 GB) governs, far above any legitimate media/model/image size.
    /// Both <see cref="ImageSource"/> and <see cref="FileSource"/> read through
    /// <see cref="ReadBounded"/> so the limit can never drift apart between them
    /// (same rationale as the SSRF policy living here).
    /// </summary>
    public const long MaxRemoteBytes = 100L * 1024 * 1024; // 100 MB

    /// <summary>
    /// Copy <paramref name="src"/> into memory, refusing once <paramref name="max"/>
    /// bytes have been read. Use after <see cref="CreateGuardedHandler"/> so a
    /// chunked / Content-Length-lying response cannot exhaust memory. Callers
    /// should still pre-check <c>response.Content.Headers.ContentLength</c> to
    /// fail fast when the server is honest about an oversized body.
    /// </summary>
    /// <param name="what">Noun used in the refusal message, e.g. "image" or "file".</param>
    public static byte[] ReadBounded(Stream src, long max, string url, string what = "file")
    {
        using var ms = new MemoryStream();
        var buf = new byte[81920];
        long total = 0;
        int n;
        while ((n = src.Read(buf, 0, buf.Length)) > 0)
        {
            total += n;
            if (total > max)
                throw new ArgumentException(
                    $"Remote {what} at {url} exceeds {max / (1024 * 1024)} MB limit.");
            ms.Write(buf, 0, n);
        }
        return ms.ToArray();
    }

    /// <summary>
    /// True only for globally-routable addresses. Blocks loopback, private
    /// (RFC1918), link-local (incl. 169.254.0.0/16 cloud-metadata), unique-local
    /// IPv6 (fc00::/7), multicast and unspecified — the SSRF target ranges.
    /// </summary>
    public static bool IsPublicAddress(IPAddress address)
    {
        var addr = address.IsIPv4MappedToIPv6 ? address.MapToIPv4() : address;

        if (IPAddress.IsLoopback(addr)) return false;
        if (addr.Equals(IPAddress.Any) || addr.Equals(IPAddress.IPv6Any)) return false;

        if (addr.AddressFamily == AddressFamily.InterNetwork)
        {
            var b = addr.GetAddressBytes(); // big-endian
            if (b[0] == 10) return false;                                   // 10.0.0.0/8
            if (b[0] == 172 && b[1] >= 16 && b[1] <= 31) return false;      // 172.16.0.0/12
            if (b[0] == 192 && b[1] == 168) return false;                   // 192.168.0.0/16
            if (b[0] == 169 && b[1] == 254) return false;                   // 169.254.0.0/16 link-local (cloud metadata)
            if (b[0] == 127) return false;                                  // 127.0.0.0/8
            if (b[0] == 100 && b[1] >= 64 && b[1] <= 127) return false;     // 100.64.0.0/10 CGNAT
            if (b[0] == 0) return false;                                    // 0.0.0.0/8
            if (b[0] >= 224) return false;                                  // 224.0.0.0/4 multicast + 240/4 reserved
            return true;
        }

        if (addr.AddressFamily == AddressFamily.InterNetworkV6)
        {
            if (addr.IsIPv6LinkLocal || addr.IsIPv6SiteLocal || addr.IsIPv6Multicast) return false;
            var b = addr.GetAddressBytes();
            if ((b[0] & 0xFE) == 0xFC) return false;                        // fc00::/7 unique-local
            return true;
        }

        return false;
    }
}
