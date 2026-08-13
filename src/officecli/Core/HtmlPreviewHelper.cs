// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeCli.Core;

/// <summary>
/// Shared helpers for HTML preview rendering across PowerPoint, Word, and Excel handlers.
/// </summary>
internal static class HtmlPreviewHelper
{
    /// <summary>
    /// HTML-encode text for safe insertion into element content or double-quoted
    /// attribute values: escapes &amp;, &lt;, &gt;, double-quote, and single-quote.
    /// This is the plain entity-encoding shared by the PowerPoint, Excel, and chart
    /// SVG renderers. (Word's preview uses a variant that additionally preserves
    /// consecutive spaces as non-breaking spaces and does not escape the apostrophe —
    /// see WordHandler.HtmlPreview.Css.HtmlEncode, kept separate by design.)
    /// </summary>
    public static string HtmlEncode(string text)
    {
        return text
            .Replace("&", "&amp;")
            .Replace("<", "&lt;")
            .Replace(">", "&gt;")
            .Replace("\"", "&quot;")
            .Replace("'", "&#39;");
    }

    /// <summary>
    /// Load an OpenXML part by its relationship ID and return the content as a base64 data URI.
    /// Returns null if the part cannot be found or read.
    /// </summary>
    public static string? PartToDataUri(OpenXmlPart parentPart, string relId)
    {
        try
        {
            var part = parentPart.GetPartById(relId);
            using var stream = part.GetStream();
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            var contentType = part.ContentType ?? "image/png";
            var undecodableLabel = UndecodableImageLabel(contentType);
            if (undecodableLabel != null)
            {
                // WMF/EMF metafiles and TIFF rasters cannot be rendered by browsers in an
                // <img> tag, and there is no cross-platform .NET rasterizer/transcoder
                // available (the project deliberately avoids System.Drawing/GDI). Degrade
                // gracefully to a self-contained SVG placeholder so the preview shows a
                // clean framed box instead of a broken-image icon.
                return PlaceholderDataUri(undecodableLabel);
            }
            return $"data:{contentType};base64,{Convert.ToBase64String(ms.ToArray())}";
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Returns a short placeholder label (e.g. "WMF", "EMF", "TIFF") for image content
    /// types that browsers cannot render in an &lt;img&gt; tag, or null for natively
    /// renderable rasters (PNG/JPG/GIF/BMP/WebP) and SVG. Covers:
    ///   - WMF/EMF metafiles: image/wmf, image/x-wmf, image/emf, image/x-emf
    ///   - TIFF rasters: image/tiff, image/tif, image/x-tiff
    /// (no cross-platform decoder/transcoder is available, so even though TIFF is a
    /// raster format it degrades to a placeholder like the metafiles).
    /// </summary>
    private static string? UndecodableImageLabel(string contentType)
    {
        if (contentType.IndexOf("emf", StringComparison.OrdinalIgnoreCase) >= 0)
            return "EMF";
        if (contentType.IndexOf("wmf", StringComparison.OrdinalIgnoreCase) >= 0)
            return "WMF";
        if (contentType.IndexOf("tif", StringComparison.OrdinalIgnoreCase) >= 0)
            return "TIFF";
        return null;
    }

    /// <summary>
    /// Build a base64-encoded SVG data URI placeholder for an undecodable image.
    /// The SVG uses a viewBox + preserveAspectRatio so it scales to fill the host
    /// &lt;img&gt; width/height, drawing a light-gray bordered rectangle with a centered
    /// label (WMF/EMF/TIFF). Base64 encoding avoids any data-URI escaping concerns.
    /// </summary>
    private static string PlaceholderDataUri(string label)
    {
        var svg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 120 120\" " +
            "preserveAspectRatio=\"xMidYMid meet\">" +
            "<rect x=\"2\" y=\"2\" width=\"116\" height=\"116\" rx=\"4\" " +
            "fill=\"#f5f5f5\" stroke=\"#cccccc\" stroke-width=\"2\"/>" +
            "<text x=\"60\" y=\"66\" font-family=\"sans-serif\" font-size=\"22\" " +
            "fill=\"#999999\" text-anchor=\"middle\">" + label + "</text>" +
            "</svg>";
        var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));
        return $"data:image/svg+xml;base64,{b64}";
    }
}
