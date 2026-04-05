/*
 * CurveShockFunctions.cs — Apply a series of triangular shocks to a stored yield curve.
 *
 * A triangular shock centred at T_c with half-width W and amplitude A adds:
 *
 *   shock(t) = A * max(0, 1 - |t - T_c| / W)
 *
 * to the continuously-compounded zero rate at every point on the base curve.
 * Multiple shocks are summed.  The result is stored as a new ZeroCurve handle
 * in the object cache.
 *
 * Typical workflow:
 *   A1: =QL_BuildSwapCurve(tenors, rates, TODAY())
 *   A2: =QL_ApplyTriangularShocks(A1, {2,5,10}, {0.001,-0.002,0.0015}, {1,2,1.5}, TODAY())
 *   A3: =QL_FixedBondPrice(100, 0.05, issueDate, matDate, A2, TODAY())
 */

using System;
using System.Text;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.TermStructures
{
    public static class CurveShockFunctions
    {
        // Default tenor grid in years used when the caller omits gridTenors.
        // Covers money-market out to 30Y with denser sampling in the short end.
        private static readonly double[] DefaultGrid =
        {
            1.0/12, 2.0/12, 3.0/12, 6.0/12, 9.0/12,
            1, 2, 3, 4, 5, 6, 7, 8, 9, 10,
            12, 15, 20, 25, 30
        };

        // ─── Shock maths ──────────────────────────────────────────────────────────

        /// <summary>
        /// Value of one triangular shock at time <paramref name="t"/> years.
        /// Returns 0 outside [center−halfWidth, center+halfWidth].
        /// </summary>
        private static double TriShock(double t, double center, double amplitude, double halfWidth)
        {
            double dist = Math.Abs(t - center);
            return dist >= halfWidth ? 0.0 : amplitude * (1.0 - dist / halfWidth);
        }

        /// <summary>Sum of all shocks at time <paramref name="t"/>.</summary>
        private static double TotalShock(double t, double[] centers, double[] amplitudes, double[] halfWidths)
        {
            double sum = 0.0;
            for (int i = 0; i < centers.Length; i++)
                sum += TriShock(t, centers[i], amplitudes[i], halfWidths[i]);
            return sum;
        }

        // ─── Cache key ────────────────────────────────────────────────────────────

        private static string ShockedCurveKey(string baseHandle, double evalDate,
                                               double[] centers, double[] amplitudes,
                                               double[] halfWidths, double[] grid)
        {
            var sb = new StringBuilder("SHCK|");
            sb.Append(baseHandle);   sb.Append('|');
            sb.Append((long)evalDate); sb.Append('|');
            sb.Append(string.Join(",", centers));    sb.Append('|');
            sb.Append(string.Join(",", amplitudes)); sb.Append('|');
            sb.Append(string.Join(",", halfWidths)); sb.Append('|');
            sb.Append(string.Join(",", grid));
            return sb.ToString();
        }

        // ─── Helpers ──────────────────────────────────────────────────────────────

        private static bool IsMissing(object o) =>
            o is ExcelMissing || o is ExcelEmpty || o == null;

        private static double[] ResolveGrid(object gridTenors) =>
            IsMissing(gridTenors) ? DefaultGrid : QLHelper.ToDoubleArray(gridTenors);

        // ─── Public Excel functions ───────────────────────────────────────────────

        [ExcelFunction(Name = "QL_ApplyTriangularShocks",
                       Category = "QuantLib — Curve Handles",
                       Description = "Apply triangular shocks to a base yield curve and cache the result. " +
                                     "Each shock peaks at shockCenters[i] with amplitude shockAmplitudes[i] " +
                                     "and tapers to zero at ±shockHalfWidths[i] years. Shocks are added to " +
                                     "continuously-compounded zero rates on a sampled tenor grid. " +
                                     "Returns a new curve handle for use with QL_FixedBondPrice etc.")]
        public static object QL_ApplyTriangularShocks(
            [ExcelArgument(Description = "Base curve handle (from QL_BuildSwapCurve, QL_BuildFlatCurve, etc.)")] string baseHandle,
            [ExcelArgument(Description = "Center tenors of shocks in years (e.g. {2,5,10})")] object shockCenters,
            [ExcelArgument(Description = "Peak amplitude of each shock, decimal (e.g. 0.001 = +10bp)")] object shockAmplitudes,
            [ExcelArgument(Description = "Half-width of each shock in years (shock = 0 outside ±halfWidth from center)")] object shockHalfWidths,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Optional: tenor grid in years. Default: 1M,2M,...,9M,1Y,...,10Y,12Y,15Y,20Y,25Y,30Y.")] object gridTenors = null)
        {
            try
            {
                var centers    = QLHelper.ToDoubleArray(shockCenters);
                var amplitudes = QLHelper.ToDoubleArray(shockAmplitudes);
                var halfWidths = QLHelper.ToDoubleArray(shockHalfWidths);

                if (centers.Length != amplitudes.Length || centers.Length != halfWidths.Length)
                    throw new ArgumentException(
                        "shockCenters, shockAmplitudes, and shockHalfWidths must have the same number of elements.");

                if (Array.Exists(halfWidths, w => w <= 0))
                    throw new ArgumentException("All shockHalfWidths must be positive.");

                var grid = ResolveGrid(gridTenors);
                string key = ShockedCurveKey(baseHandle, evalDate, centers, amplitudes, halfWidths, grid);

                if (!ObjectCache.HasCurve(key))
                {
                    var baseCurve = ObjectCache.GetCurve(baseHandle);
                    var evalQL    = QLHelper.ToQLDate(evalDate);
                    Settings.instance().setEvaluationDate(evalQL);

                    var dc       = new Actual365Fixed();
                    var calendar = new TARGET();

                    var qlDates = new DateVector();
                    var qlRates = new DoubleVector();

                    foreach (double tenor in grid)
                    {
                        // Advance by the equivalent number of months, adjusted to business days.
                        var months   = (int)Math.Round(tenor * 12);
                        if (months < 1) months = 1;             // guard against sub-monthly grid entries
                        var pillarQL = calendar.advance(evalQL, months, TimeUnit.Months,
                                                        BusinessDayConvention.ModifiedFollowing);

                        // Base continuously-compounded zero rate at this pillar.
                        double baseRate = baseCurve.currentLink()
                                            .zeroRate(pillarQL, dc,
                                                      Compounding.Continuous, Frequency.Annual)
                                            .rate();

                        qlDates.Add(pillarQL);
                        qlRates.Add(baseRate + TotalShock(tenor, centers, amplitudes, halfWidths));
                    }

                    // Build a linearly-interpolated zero curve with flat extrapolation.
                    var shockedCurve = new ZeroCurve(qlDates, qlRates, dc, calendar,
                                                     new Linear(),
                                                     Compounding.Continuous, Frequency.Annual);
                    shockedCurve.enableExtrapolation();
                    ObjectCache.StoreCurve(key, new YieldTermStructureHandle(shockedCurve));
                }

                return key;
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_TriangularShockProfile",
                       Category = "QuantLib — Curve Handles",
                       Description = "Diagnostic: show the total shock (in decimal) at each tenor on the grid. " +
                                     "Returns an N×2 array [Tenor (years), Total Shock]. Enter as an array formula.")]
        public static object QL_TriangularShockProfile(
            [ExcelArgument(Description = "Center tenors of shocks in years")] object shockCenters,
            [ExcelArgument(Description = "Peak amplitude of each shock (decimal)")] object shockAmplitudes,
            [ExcelArgument(Description = "Half-width of each shock in years")] object shockHalfWidths,
            [ExcelArgument(Description = "Optional: tenor grid in years. Default: 1M to 30Y.")] object gridTenors = null)
        {
            try
            {
                var centers    = QLHelper.ToDoubleArray(shockCenters);
                var amplitudes = QLHelper.ToDoubleArray(shockAmplitudes);
                var halfWidths = QLHelper.ToDoubleArray(shockHalfWidths);

                if (centers.Length != amplitudes.Length || centers.Length != halfWidths.Length)
                    throw new ArgumentException(
                        "shockCenters, shockAmplitudes, and shockHalfWidths must have the same number of elements.");

                var grid   = ResolveGrid(gridTenors);
                var result = new object[grid.Length, 2];

                for (int j = 0; j < grid.Length; j++)
                {
                    result[j, 0] = grid[j];
                    result[j, 1] = TotalShock(grid[j], centers, amplitudes, halfWidths);
                }

                return result;
            }
            catch (Exception ex) { return new object[,] { { $"#QL_ERR: {ex.Message}" } }; }
        }
    }
}
