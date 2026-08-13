// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

internal partial class FormulaEvaluator
{
    // ==================== Descriptive & regression statistics ====================
    //
    // Shape (SKEW/KURT/AVEDEV/DEVSQ/TRIMMEAN) and the pairwise regression family
    // (CORREL/COVAR/SLOPE/INTERCEPT/RSQ/STEYX/FORECAST) plus the exclusive
    // quartile/percentile variants. Pure array arithmetic — no special functions.

    private double[] Flat(List<object> args) => FlattenNumbers(args);

    // Pull two aligned numeric arrays (regression takes two ranges).
    private static bool Pair(List<object> args, out double[] a, out double[] b)
    {
        a = AsDoubles(args.Count > 0 ? args[0] : null) ?? [];
        b = AsDoubles(args.Count > 1 ? args[1] : null) ?? [];
        if (a.Length == 0 || a.Length != b.Length) { a = b = []; return false; }
        return true;
    }

    private FormulaResult? EvalSkew(List<object> args, bool population)
    {
        var v = Flat(args);
        int n = v.Length;
        if (population ? n < 1 : n < 3) return FormulaResult.Error("#DIV/0!");
        double mean = v.Average();
        double sd = Math.Sqrt(v.Sum(x => (x - mean) * (x - mean)) / (population ? n : n - 1));
        if (sd == 0) return FormulaResult.Error("#DIV/0!");
        double s3 = v.Sum(x => Math.Pow((x - mean) / sd, 3));
        return FR(population ? s3 / n : (double)n / ((n - 1.0) * (n - 2.0)) * s3);
    }

    private FormulaResult? EvalKurt(List<object> args)
    {
        var v = Flat(args);
        int n = v.Length;
        if (n < 4) return FormulaResult.Error("#DIV/0!");
        double mean = v.Average();
        double sd = Math.Sqrt(v.Sum(x => (x - mean) * (x - mean)) / (n - 1.0));
        if (sd == 0) return FormulaResult.Error("#DIV/0!");
        double s4 = v.Sum(x => Math.Pow((x - mean) / sd, 4));
        return FR((double)n * (n + 1) / ((n - 1.0) * (n - 2) * (n - 3)) * s4
            - 3.0 * (n - 1) * (n - 1) / ((n - 2.0) * (n - 3)));
    }

    private FormulaResult? EvalTrimMean(List<object> args)
    {
        if (args.Count < 2) return null;
        var v = AsDoubles(args[0]); double pct = args[1] is FormulaResult r ? r.AsNumber() : 0;
        if (v == null || v.Length == 0 || pct < 0 || pct >= 1) return FormulaResult.Error("#NUM!");
        var s = v.OrderBy(x => x).ToArray();
        int trim = (int)Math.Floor(s.Length * pct / 2);   // trimmed from EACH end
        var kept = s.Skip(trim).Take(s.Length - 2 * trim).ToArray();
        return kept.Length > 0 ? FR(kept.Average()) : FormulaResult.Error("#NUM!");
    }

    private FormulaResult? EvalCorrel(List<object> args)
    {
        if (!Pair(args, out var x, out var y)) return FormulaResult.Error("#N/A");
        double mx = x.Average(), my = y.Average();
        double sxy = 0, sxx = 0, syy = 0;
        for (int i = 0; i < x.Length; i++) { double dx = x[i] - mx, dy = y[i] - my; sxy += dx * dy; sxx += dx * dx; syy += dy * dy; }
        return sxx == 0 || syy == 0 ? FormulaResult.Error("#DIV/0!") : FR(sxy / Math.Sqrt(sxx * syy));
    }

    private FormulaResult? EvalCovar(List<object> args, bool sample)
    {
        if (!Pair(args, out var x, out var y)) return FormulaResult.Error("#N/A");
        int n = x.Length;
        if (sample && n < 2) return FormulaResult.Error("#DIV/0!");
        double mx = x.Average(), my = y.Average(), s = 0;
        for (int i = 0; i < n; i++) s += (x[i] - mx) * (y[i] - my);
        return FR(s / (sample ? n - 1 : n));
    }

    // SLOPE(known_y, known_x).
    private FormulaResult? EvalSlope(List<object> args)
    {
        if (!Pair(args, out var y, out var x)) return FormulaResult.Error("#N/A");
        double mx = x.Average(), my = y.Average(), sxy = 0, sxx = 0;
        for (int i = 0; i < x.Length; i++) { double dx = x[i] - mx; sxy += dx * (y[i] - my); sxx += dx * dx; }
        return sxx == 0 ? FormulaResult.Error("#DIV/0!") : FR(sxy / sxx);
    }

    private FormulaResult? EvalIntercept(List<object> args)
    {
        if (!Pair(args, out var y, out var x)) return FormulaResult.Error("#N/A");
        double mx = x.Average(), my = y.Average(), sxy = 0, sxx = 0;
        for (int i = 0; i < x.Length; i++) { double dx = x[i] - mx; sxy += dx * (y[i] - my); sxx += dx * dx; }
        return sxx == 0 ? FormulaResult.Error("#DIV/0!") : FR(my - sxy / sxx * mx);
    }

    private FormulaResult? EvalRsq(List<object> args)
    {
        var c = EvalCorrel(args);
        return c is { IsNumeric: true } ? FR(c.NumericValue!.Value * c.NumericValue.Value) : c;
    }

    // STEYX(known_y, known_x) — standard error of the regression estimate.
    private FormulaResult? EvalSteyx(List<object> args)
    {
        if (!Pair(args, out var y, out var x)) return FormulaResult.Error("#N/A");
        int n = x.Length;
        if (n < 3) return FormulaResult.Error("#DIV/0!");
        double mx = x.Average(), my = y.Average(), sxx = 0, syy = 0, sxy = 0;
        for (int i = 0; i < n; i++) { double dx = x[i] - mx, dy = y[i] - my; sxx += dx * dx; syy += dy * dy; sxy += dx * dy; }
        return sxx == 0 ? FormulaResult.Error("#DIV/0!") : FR(Math.Sqrt((syy - sxy * sxy / sxx) / (n - 2)));
    }

    // FORECAST(x, known_y, known_x) / FORECAST.LINEAR.
    private FormulaResult? EvalForecast(List<object> args)
    {
        if (args.Count < 3) return null;
        double x0 = args[0] is FormulaResult r ? r.AsNumber() : 0;
        var pair = new List<object> { args[1], args[2] };
        if (!Pair(pair, out var y, out var x)) return FormulaResult.Error("#N/A");
        double mx = x.Average(), my = y.Average(), sxy = 0, sxx = 0;
        for (int i = 0; i < x.Length; i++) { double dx = x[i] - mx; sxy += dx * (y[i] - my); sxx += dx * dx; }
        if (sxx == 0) return FormulaResult.Error("#DIV/0!");
        double b = sxy / sxx;
        return FR(my - b * mx + b * x0);
    }

    // QUARTILE(range, quart) inclusive / exclusive — reuse the percentile engine.
    private FormulaResult? EvalQuartile(List<object> args, bool exclusive)
    {
        if (args.Count < 2) return null;
        int q = args[1] is FormulaResult r ? (int)r.AsNumber() : 0;
        if (q is < 0 or > 4) return FormulaResult.Error("#NUM!");
        return PercentileAt(AsDoubles(args[0]), q / 4.0, exclusive);
    }

    private FormulaResult? EvalPercentileExc(List<object> args)
    {
        if (args.Count < 2) return null;
        double k = args[1] is FormulaResult r ? r.AsNumber() : 0;
        return PercentileAt(AsDoubles(args[0]), k, exclusive: true);
    }

    // Shared percentile interpolation. Inclusive: rank = k(n-1). Exclusive:
    // rank = k(n+1)-1, valid only for 1/(n+1) ≤ k ≤ n/(n+1).
    private static FormulaResult? PercentileAt(double[]? data, double k, bool exclusive)
    {
        if (data == null || data.Length == 0) return FormulaResult.Error("#NUM!");
        var s = data.OrderBy(x => x).ToArray();
        int n = s.Length;
        double pos = exclusive ? k * (n + 1) - 1 : k * (n - 1);
        if (exclusive && (pos < 0 || pos > n - 1)) return FormulaResult.Error("#NUM!");
        if (!exclusive && (k < 0 || k > 1)) return FormulaResult.Error("#NUM!");
        int lo = (int)Math.Floor(pos);
        double frac = pos - lo;
        if (lo + 1 < n) return FR(s[lo] + frac * (s[lo + 1] - s[lo]));
        return FR(s[lo]);
    }
}
