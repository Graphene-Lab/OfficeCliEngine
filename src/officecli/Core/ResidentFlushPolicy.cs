using System;

namespace OfficeCli.Core;

/// <summary>
/// Resident flush policy: when the resident's in-memory DOM is written to disk.
/// One knob with four modes (OFFICECLI_RESIDENT_FLUSH, legacy alias
/// OFFICECLI_RESIDENT_IDLE_SAVE_SECONDS):
///   each  — flush before every mutation command returns (deterministic; the
///           caller pays one O(n) serialize per mutation, batch still one at end)
///   auto  — idle-debounced flush with an adaptive interval derived from
///           measured save durations (default)
///   &lt;N&gt;   — idle-debounced flush with a fixed N-second interval
///   off/0 — never auto-flush; only explicit save/close/shutdown write to disk
/// </summary>
internal enum ResidentFlushMode
{
    Each,
    Auto,
    Fixed,
    Off,
}

public static class ResidentFlushPolicy
{
    /// <summary>
    /// Adaptive interval = clamp(CostMultiplier × EMA(save duration), Min, Max).
    /// The multiplier bounds the background save's share of wall-clock time at
    /// 1/CostMultiplier even for data-heavy files whose save takes seconds, so
    /// a shorter debounce can never degenerate into a save storm.
    /// </summary>
    internal const double CostMultiplier = 4.0;
    internal static readonly TimeSpan MinAdaptiveInterval = TimeSpan.FromSeconds(2);
    internal static readonly TimeSpan MaxAdaptiveInterval = TimeSpan.FromSeconds(10);

    // Asymmetric smoothing: a slow save raises the estimate almost immediately
    // (protect against save storms right away); a fast save lowers it only
    // gradually (a single quick save — or the post-cold-start drop — must not
    // whipsaw the interval back down).
    internal const double RiseAlpha = 0.7;
    internal const double FallAlpha = 0.2;

    /// <summary>
    /// Fold one measured save duration into the EMA. A negative previous value
    /// means "no sample yet": the first measurement seeds the EMA directly.
    /// Public so embedders reuse the same debounce
    /// decision instead of growing a second, driftable copy.
    /// </summary>
    public static double NextEmaSeconds(double previousEmaSeconds, double sampleSeconds)
    {
        if (sampleSeconds < 0) sampleSeconds = 0;
        if (previousEmaSeconds < 0) return sampleSeconds;
        var alpha = sampleSeconds > previousEmaSeconds ? RiseAlpha : FallAlpha;
        return alpha * sampleSeconds + (1 - alpha) * previousEmaSeconds;
    }

    /// <summary>
    /// Debounce interval for the current EMA. No sample yet (negative EMA) →
    /// the floor: typical documents save in well under Min/CostMultiplier, and
    /// a data-heavy file pays at most one early save before the EMA raises it.
    /// Public — see <see cref="NextEmaSeconds"/>.
    /// </summary>
    public static TimeSpan IntervalForEma(double emaSeconds)
    {
        if (emaSeconds < 0) return MinAdaptiveInterval;
        var seconds = CostMultiplier * emaSeconds;
        if (seconds < MinAdaptiveInterval.TotalSeconds) return MinAdaptiveInterval;
        if (seconds > MaxAdaptiveInterval.TotalSeconds) return MaxAdaptiveInterval;
        return TimeSpan.FromSeconds(seconds);
    }

    /// <summary>
    /// Parse a policy string: "each" | "auto" | "off"/"0"/non-positive int |
    /// integer seconds (bounded by [minSeconds, maxSeconds]). Returns false for
    /// unrecognized or out-of-range input (caller falls back to the default).
    /// </summary>
    internal static bool TryParse(string? raw, int minSeconds, int maxSeconds,
        out ResidentFlushMode mode, out TimeSpan fixedInterval)
    {
        mode = ResidentFlushMode.Auto;
        fixedInterval = default;
        if (string.IsNullOrWhiteSpace(raw)) return false;
        var value = raw.Trim();
        if (value.Equals("each", StringComparison.OrdinalIgnoreCase))
        {
            mode = ResidentFlushMode.Each;
            return true;
        }
        if (value.Equals("auto", StringComparison.OrdinalIgnoreCase))
        {
            mode = ResidentFlushMode.Auto;
            return true;
        }
        if (value.Equals("off", StringComparison.OrdinalIgnoreCase))
        {
            mode = ResidentFlushMode.Off;
            return true;
        }
        if (int.TryParse(value, out var secs))
        {
            if (secs <= 0)
            {
                mode = ResidentFlushMode.Off;
                return true;
            }
            if (secs >= minSeconds && secs <= maxSeconds)
            {
                mode = ResidentFlushMode.Fixed;
                fixedInterval = TimeSpan.FromSeconds(secs);
                return true;
            }
        }
        return false;
    }
}
