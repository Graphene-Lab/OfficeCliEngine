// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

internal partial class FormulaEvaluator
{
    // ==================== Iterative root solver ====================
    //
    // Shared foundation for every Excel function whose value is the root of an
    // equation rather than a closed form: RATE / IRR / XIRR / MIRR-style yield,
    // and later YIELD / ODDFYIELD / etc. One solver, mirrored on Excel's
    // approach: seeded Newton, bounded iterations, #NUM! on non-convergence.
    //
    // Newton with a numerical derivative is the primary method (matches Excel's
    // quadratic convergence on smooth NPV curves); a bisection fallback over an
    // expanding bracket recovers the cases where Newton's derivative vanishes or
    // the iterate escapes the convergence basin (multiple sign changes, deep
    // out-of-the-money cashflows). Returns null when neither converges — callers
    // surface that as #NUM!, exactly like Excel.

    private const int SolverMaxIter = 100;
    private const double SolverTol = 1e-10;

    /// <summary>
    /// Find x such that f(x) ≈ 0, seeded at <paramref name="guess"/>. Newton
    /// first (numerical derivative), then an expanding-bracket bisection fallback
    /// clamped to the rate domain (-1, ∞). Null = no root found → caller emits
    /// #NUM!. The fallback matters: real Excel's IRR/RATE recover far in-domain
    /// roots that naive bounded Newton overshoots (e.g. a never-recoups series
    /// solving to ≈ -42%), so the fallback is what keeps us equal to Excel's
    /// cached value rather than a spreadsheet engine that gives up there.
    /// </summary>
    private static double? SolveRoot(Func<double, double> f, double guess)
    {
        double x = guess;
        for (int i = 0; i < SolverMaxIter; i++)
        {
            double fx = f(x);
            if (double.IsNaN(fx) || double.IsInfinity(fx)) break;
            if (Math.Abs(fx) < SolverTol) return x;

            // Central numerical derivative; step scales with |x| for conditioning.
            double h = Math.Max(1e-7, Math.Abs(x) * 1e-7);
            double dfx = (f(x + h) - f(x - h)) / (2 * h);
            if (double.IsNaN(dfx) || double.IsInfinity(dfx) || Math.Abs(dfx) < 1e-14) break;

            double next = x - fx / dfx;
            if (double.IsNaN(next) || double.IsInfinity(next)) break;
            if (Math.Abs(next - x) < SolverTol) return next;
            x = next;
        }
        return BisectFallback(f, guess);
    }

    /// <summary>Expand a bracket outward from the guess until f changes sign,
    /// then bisect. Covers rate-style roots in (-1, ∞).</summary>
    private static double? BisectFallback(Func<double, double> f, double guess)
    {
        // Rates live in (-1, ∞); clamp the lower probe just above -1 so
        // (1+rate)^n stays defined.
        double lo = -0.999999, hi = Math.Max(1.0, guess + 1.0);
        double flo = f(lo);
        if (double.IsNaN(flo) || double.IsInfinity(flo)) return null;

        double fhi = f(hi);
        int expand = 0;
        while ((double.IsNaN(fhi) || double.IsInfinity(fhi) || Math.Sign(flo) == Math.Sign(fhi)) && expand < 200)
        {
            hi += Math.Max(1.0, Math.Abs(hi));   // grow the upper bound
            fhi = f(hi);
            expand++;
        }
        if (double.IsNaN(fhi) || double.IsInfinity(fhi) || Math.Sign(flo) == Math.Sign(fhi))
            return null;   // no sign change found — genuinely no root in range

        for (int i = 0; i < SolverMaxIter * 2; i++)
        {
            double mid = (lo + hi) / 2;
            double fmid = f(mid);
            if (double.IsNaN(fmid) || double.IsInfinity(fmid)) return null;
            if (Math.Abs(fmid) < SolverTol || (hi - lo) / 2 < SolverTol) return mid;
            if (Math.Sign(fmid) == Math.Sign(flo)) { lo = mid; flo = fmid; }
            else { hi = mid; }
        }
        return (lo + hi) / 2;
    }
}
