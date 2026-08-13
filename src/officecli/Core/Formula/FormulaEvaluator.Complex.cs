// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Globalization;
using System.Numerics;

namespace OfficeCli.Core;

internal partial class FormulaEvaluator
{
    // ==================== Complex number foundation ====================
    //
    // Excel stores complex numbers as TEXT ("3+4i" / "2-3j"). This is the shared
    // base for the whole IM* family: parse an Excel complex string into a
    // System.Numerics.Complex, run the math there, and format the result back to
    // Excel's exact textual form (coefficient-1 omitted, suffix preserved). The
    // engineering functions are then thin wrappers over Complex's operators.

    // Parse "a+bi" / "a-bj" / "bi" / "a" / "i" / "-i" into a Complex. Returns
    // false on a malformed string (caller surfaces #NUM!). suffix echoes the
    // imaginary unit found ('i' default, 'j' if the input used j).
    private static bool TryParseComplex(string s, out Complex value, out char suffix)
    {
        value = Complex.Zero;
        suffix = 'i';
        s = s.Trim();
        if (s.Length == 0) { value = Complex.Zero; return true; }

        char last = s[^1];
        if (last is 'i' or 'j' or 'I' or 'J')
        {
            suffix = char.ToLowerInvariant(last);
            string body = s[..^1];      // strip the imaginary unit

            // Split real and imaginary at the +/- that separates them — not the
            // leading sign and not the sign of an exponent (e+3).
            int split = -1;
            for (int k = 1; k < body.Length; k++)
                if (body[k] is '+' or '-' && body[k - 1] is not ('e' or 'E'))
                    split = k;

            string realStr = split < 0 ? "0" : body[..split];
            string imagStr = split < 0 ? body : body[split..];
            // empty / lone-sign imaginary coefficient means ±1 ("i", "-i", "3+i")
            double imag = imagStr is "" or "+" ? 1 : imagStr is "-" ? -1
                : (TryNum(imagStr, out var iv) ? iv : double.NaN);
            double real = realStr is "" or "+" ? 0 : realStr is "-" ? 0
                : (TryNum(realStr, out var rv) ? rv : double.NaN);
            if (double.IsNaN(real) || double.IsNaN(imag)) return false;
            value = new Complex(real, imag);
            return true;
        }
        // purely real
        if (!TryNum(s, out var realOnly)) return false;
        value = new Complex(realOnly, 0);
        return true;

        static bool TryNum(string x, out double d) =>
            double.TryParse(x, NumberStyles.Any, CultureInfo.InvariantCulture, out d);
    }

    // Format a Complex the way Excel does: "3+4i", "3-4i", "5", "i", "-i", "2i".
    private static string FormatComplex(Complex c, char suffix)
    {
        double re = c.Real, im = c.Imaginary;
        // Clean up -0 so it never prints a stray minus.
        if (re == 0) re = 0;
        if (im == 0) im = 0;

        if (im == 0) return FmtNum(re);
        string imagPart = im switch
        {
            1 => $"{suffix}",
            -1 => $"-{suffix}",
            _ => $"{FmtNum(im)}{suffix}",
        };
        if (re == 0) return imagPart;

        // both parts present — join with the imaginary part's sign
        if (im > 0)
            return im == 1 ? $"{FmtNum(re)}+{suffix}" : $"{FmtNum(re)}+{FmtNum(im)}{suffix}";
        return im == -1 ? $"{FmtNum(re)}-{suffix}" : $"{FmtNum(re)}-{FmtNum(Math.Abs(im))}{suffix}";
    }

    // Excel general-number text: up to 15 significant digits, trailing zeros
    // trimmed. Matches the cell renderer's 15-sig rounding so a complex string's
    // components line up with what Excel persists.
    private static string FmtNum(double v)
    {
        if (v == 0) return "0";
        if (double.IsNaN(v) || double.IsInfinity(v)) return "#NUM!";
        // G15 = exactly Excel's 15-significant-digit width (matches the stored
        // complex string component); "R" would emit up to 17 round-trip digits
        // and diverge in the last place.
        return v.ToString("G15", CultureInfo.InvariantCulture);
    }

    // Pull a complex argument out of the arg list (string or bare number).
    private static bool ArgComplex(object? a, out Complex c, out char suffix)
    {
        c = Complex.Zero; suffix = 'i';
        if (a is FormulaResult { IsNumeric: true } n) { c = new Complex(n.NumericValue!.Value, 0); return true; }
        if (a is FormulaResult r) return TryParseComplex(r.AsString(), out c, out suffix);
        return false;
    }

    // COMPLEX(real, imaginary, [suffix]).
    private FormulaResult? EvalComplex(List<object> args)
    {
        if (args.Count < 2) return null;
        double re = args[0] is FormulaResult r ? r.AsNumber() : 0;
        double im = args[1] is FormulaResult i ? i.AsNumber() : 0;
        char suf = 'i';
        if (args.Count > 2 && args[2] is FormulaResult sfx)
        {
            var ss = sfx.AsString();
            if (ss is not ("i" or "j")) return FormulaResult.Error("#VALUE!");
            suf = ss[0];
        }
        return FR_S(FormatComplex(new Complex(re, im), suf));
    }

    // Unary IM* returning a complex string (IMSQRT, IMEXP, IMSIN, ...).
    private FormulaResult? ImUnary(List<object> args, Func<Complex, Complex> f)
    {
        if (args.Count < 1 || !ArgComplex(args[0], out var c, out var suf)) return FormulaResult.Error("#NUM!");
        var result = f(c);
        if (double.IsNaN(result.Real) || double.IsNaN(result.Imaginary)
            || double.IsInfinity(result.Real) || double.IsInfinity(result.Imaginary))
            return FormulaResult.Error("#NUM!");   // e.g. IMLN(0), IMDIV by 0
        return FR_S(FormatComplex(result, suf));
    }

    // Unary IM* returning a real number (IMABS, IMREAL, IMAGINARY, IMARGUMENT).
    private FormulaResult? ImScalar(List<object> args, Func<Complex, double> f)
    {
        if (args.Count < 1 || !ArgComplex(args[0], out var c, out _)) return FormulaResult.Error("#NUM!");
        return FR(f(c));
    }

    // Variadic IM* folding a list of complex args (IMSUM, IMPRODUCT).
    private FormulaResult? ImFold(List<object> args, Complex seed, Func<Complex, Complex, Complex> op)
    {
        Complex acc = seed; char suf = 'i'; bool any = false;
        foreach (var a in AllArgs(args))
        {
            if (!ArgComplex(a, out var c, out var s)) return FormulaResult.Error("#NUM!");
            if (any && s != suf) return FormulaResult.Error("#NUM!");   // mixed i/j
            suf = s; acc = any ? op(acc, c) : c; any = true;
        }
        return any ? FR_S(FormatComplex(acc, suf)) : FormulaResult.Error("#NUM!");
    }

    // Binary IM* (IMSUB, IMDIV).
    private FormulaResult? ImBinary(List<object> args, Func<Complex, Complex, Complex> op)
    {
        if (args.Count < 2 || !ArgComplex(args[0], out var a, out var sa) || !ArgComplex(args[1], out var b, out var sb))
            return FormulaResult.Error("#NUM!");
        if (sa != sb && a.Imaginary != 0 && b.Imaginary != 0) return FormulaResult.Error("#NUM!");
        return FR_S(FormatComplex(op(a, b), sa));
    }
}
