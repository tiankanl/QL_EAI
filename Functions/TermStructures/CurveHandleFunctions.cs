/*
 * CurveHandleFunctions.cs — Build yield curves and store them in the object cache.
 *
 * Workflow:
 *   A1: =QL_BuildSwapCurve(tenors, rates, TODAY())     → "SWCV|46921|1,2,3,5,7,10|..."
 *   B2: =QL_FixedBondCashflowsH($A$1, 100, 0.05, ...)  → cashflow table
 *   B3: =QL_FRNPriceH($A$1, 100, 0.005, ...)           → price array
 *
 *  The handle in A1 changes only when inputs change; all downstream cells
 *  recalculate only when A1 changes — the curve is never rebuilt twice per inputs.
 */

using System;
using System.Text;
using SM = System.Math;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.TermStructures
{
    public static class CurveHandleFunctions
    {
        // ─── Key generation ───────────────────────────────────────────────────────

        private static string SwapCurveKey(double evalDate, double[] tenors, double[] rates)
        {
            var sb = new StringBuilder("SWCV|");
            sb.Append((long)evalDate);
            sb.Append('|');
            sb.Append(string.Join(",", tenors));
            sb.Append('|');
            sb.Append(string.Join(",", rates));
            return sb.ToString();
        }

        private static string FlatCurveKey(double evalDate, double rate)
            => $"FLAT|{(long)evalDate}|{rate:R}";

        private static string PointCurveKey(string type, double evalDate,
                                             double[] dates, double[] values)
        {
            var sb = new StringBuilder(type); sb.Append('|');
            sb.Append((long)evalDate); sb.Append('|');
            sb.Append(string.Join(",", dates)); sb.Append('|');
            sb.Append(string.Join(",", values));
            return sb.ToString();
        }

        // ─── Swap curve builder ───────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_BuildSwapCurve",
                       Category = "QuantLib — Curve Handles",
                       Description = "Bootstrap a par-swap yield curve and store it in the in-process cache.\n" +
                                     "Returns a handle string — pass this to QL_FixedBondCashflowsH, QL_FRNPriceH, etc.\n" +
                                     "The curve is only rebuilt when inputs change; all downstream functions are free.")]
        public static object QL_BuildSwapCurve(
            [ExcelArgument(Description = "Range of tenor years (e.g. 1,2,3,5,7,10)")] object tenorYears,
            [ExcelArgument(Description = "Range of par swap rates (decimals, same order)")] object parRates,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate)
        {
            try
            {
                var tenors = QLHelper.ToDoubleArray(tenorYears);
                var rates  = QLHelper.ToDoubleArray(parRates);
                string key = SwapCurveKey(evalDate, tenors, rates);

                if (!ObjectCache.HasCurve(key))
                {
                    var evalQL   = QLHelper.ToQLDate(evalDate);
                    Settings.instance().setEvaluationDate(evalQL);

                    var calendar = new TARGET();
                    var dc       = new Actual365Fixed();
                    var fixedDC  = new Thirty360(Thirty360.Convention.BondBasis);
                    var index    = new Euribor6M();

                    var helpers = new RateHelperVector();
                    for (int i = 0; i < tenors.Length; i++)
                    {
                        var quote  = new QuoteHandle(new SimpleQuote(rates[i]));
                        var tenor  = new Period((int)SM.Round(tenors[i]), TimeUnit.Years);
                        helpers.Add(new SwapRateHelper(
                            quote, tenor, calendar,
                            Frequency.Annual,
                            BusinessDayConvention.ModifiedFollowing,
                            fixedDC, index));
                    }

                    var curve  = new PiecewiseFlatForward(evalQL, helpers, dc);
                    var handle = new YieldTermStructureHandle(curve);
                    ObjectCache.StoreCurve(key, handle);
                }

                return key;
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Flat curve builder ───────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_BuildFlatCurve",
                       Category = "QuantLib — Curve Handles",
                       Description = "Build a flat yield curve and store it in the in-process cache.\n" +
                                     "Returns a handle string for use with QL_FixedBondCashflowsH, QL_FRNPriceH, etc.")]
        public static object QL_BuildFlatCurve(
            [ExcelArgument(Description = "Flat continuously-compounded rate (decimal)")] double rate,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate)
        {
            try
            {
                string key = FlatCurveKey(evalDate, rate);

                if (!ObjectCache.HasCurve(key))
                {
                    var evalQL = QLHelper.ToQLDate(evalDate);
                    Settings.instance().setEvaluationDate(evalQL);
                    var curve  = new FlatForward(evalQL, rate, new Actual365Fixed());
                    ObjectCache.StoreCurve(key, new YieldTermStructureHandle(curve));
                }

                return key;
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Curve queries from handle ────────────────────────────────────────────

        [ExcelFunction(Name = "QL_CurveDiscount",
                       Category = "QuantLib — Curve Handles",
                       Description = "Discount factor P(0,T) from a stored curve handle.")]
        public static object QL_CurveDiscount(
            [ExcelArgument(Description = "Curve handle (from QL_BuildSwapCurve / QL_BuildFlatCurve)")] string curveHandle,
            [ExcelArgument(Description = "Query date (Excel date)")] double queryDate)
        {
            try
            {
                var handle  = ObjectCache.GetCurve(curveHandle);
                return handle.currentLink().discount(QLHelper.ToQLDate(queryDate));
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_CurveZeroRate",
                       Category = "QuantLib — Curve Handles",
                       Description = "Zero rate at a given date from a stored curve handle.")]
        public static object QL_CurveZeroRate(
            [ExcelArgument(Description = "Curve handle")] string curveHandle,
            [ExcelArgument(Description = "Query date (Excel date)")] double queryDate,
            [ExcelArgument(Description = "Compounding (default Continuous)")] string compounding = "Continuous",
            [ExcelArgument(Description = "Frequency (default Annual)")] string frequency = "Annual")
        {
            try
            {
                var handle = ObjectCache.GetCurve(curveHandle);
                var dc     = new Actual365Fixed();
                var comp   = QLHelper.ParseCompounding(compounding);
                var freq   = QLHelper.ParseFrequency(frequency);
                return handle.currentLink()
                             .zeroRate(QLHelper.ToQLDate(queryDate), dc, comp, freq).rate();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_CurveForwardRate",
                       Category = "QuantLib — Curve Handles",
                       Description = "Forward rate between two dates from a stored curve handle.")]
        public static object QL_CurveForwardRate(
            [ExcelArgument(Description = "Curve handle")] string curveHandle,
            [ExcelArgument(Description = "Forward period start date")] double date1,
            [ExcelArgument(Description = "Forward period end date")] double date2,
            [ExcelArgument(Description = "Compounding (default Continuous)")] string compounding = "Continuous",
            [ExcelArgument(Description = "Frequency (default Annual)")] string frequency = "Annual")
        {
            try
            {
                var handle = ObjectCache.GetCurve(curveHandle);
                var dc     = new Actual365Fixed();
                var comp   = QLHelper.ParseCompounding(compounding);
                var freq   = QLHelper.ParseFrequency(frequency);
                return handle.currentLink()
                             .forwardRate(QLHelper.ToQLDate(date1), QLHelper.ToQLDate(date2),
                                          dc, comp, freq).rate();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Zero-rate curve builder ──────────────────────────────────────────────

        [ExcelFunction(Name = "QL_BuildZeroCurve",
                       Category = "QuantLib — Curve Handles",
                       Description = "Build a yield curve from (date, zero rate) pairs and store it in the cache.\n" +
                                     "Interpolates linearly between zero rates; flat extrapolation beyond endpoints.\n" +
                                     "dates: range of Excel date serials. zeroRates: continuously-compounded decimals.\n" +
                                     "Returns a handle string for use with QL_FixedBondCashflows, QL_FRNPrice, etc.")]
        public static object QL_BuildZeroCurve(
            [ExcelArgument(Description = "Range of dates (Excel date serials, must include eval date as first point)")] object dates,
            [ExcelArgument(Description = "Range of continuously-compounded zero rates (decimals, same order as dates)")] object zeroRates,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365")
        {
            try
            {
                var datesArr = QLHelper.ToDoubleArray(dates);
                var ratesArr = QLHelper.ToDoubleArray(zeroRates);
                if (datesArr.Length != ratesArr.Length)
                    throw new ArgumentException("dates and zeroRates must have the same number of elements.");

                string key = PointCurveKey("ZERO", evalDate, datesArr, ratesArr);

                if (!ObjectCache.HasCurve(key))
                {
                    var evalQL = QLHelper.ToQLDate(evalDate);
                    Settings.instance().setEvaluationDate(evalQL);
                    var dc     = QLHelper.ParseDayCounter(dayCounter);

                    var qlDates = new DateVector();
                    var qlRates = new DoubleVector();
                    for (int i = 0; i < datesArr.Length; i++)
                    {
                        qlDates.Add(QLHelper.ToQLDate(datesArr[i]));
                        qlRates.Add(ratesArr[i]);
                    }

                    // ZeroCurve = InterpolatedZeroCurve<Linear> with continuous compounding.
                    var curve  = new ZeroCurve(qlDates, qlRates, dc, new TARGET(),
                                               new Linear(),
                                               Compounding.Continuous, Frequency.Annual);
                    curve.enableExtrapolation();
                    ObjectCache.StoreCurve(key, new YieldTermStructureHandle(curve));
                }

                return key;
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Discount-factor curve builder ────────────────────────────────────────

        [ExcelFunction(Name = "QL_BuildDiscountCurve",
                       Category = "QuantLib — Curve Handles",
                       Description = "Build a yield curve from (date, discount factor) pairs and store it in the cache.\n" +
                                     "Interpolates log-linearly between discount factors (equivalent to linear on zero rates).\n" +
                                     "dates: range of Excel date serials. discountFactors: decimals (e.g. 0.956 for P(0,5Y)).\n" +
                                     "Returns a handle string for use with QL_FixedBondCashflows, QL_FRNPrice, etc.")]
        public static object QL_BuildDiscountCurve(
            [ExcelArgument(Description = "Range of dates (Excel date serials)")] object dates,
            [ExcelArgument(Description = "Range of discount factors (decimals, P(0,T) ≤ 1)")] object discountFactors,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365")
        {
            try
            {
                var datesArr   = QLHelper.ToDoubleArray(dates);
                var factorsArr = QLHelper.ToDoubleArray(discountFactors);
                if (datesArr.Length != factorsArr.Length)
                    throw new ArgumentException("dates and discountFactors must have the same number of elements.");

                string key = PointCurveKey("DISC", evalDate, datesArr, factorsArr);

                if (!ObjectCache.HasCurve(key))
                {
                    var evalQL = QLHelper.ToQLDate(evalDate);
                    Settings.instance().setEvaluationDate(evalQL);
                    var dc     = QLHelper.ParseDayCounter(dayCounter);

                    var qlDates   = new DateVector();
                    var qlFactors = new DoubleVector();
                    for (int i = 0; i < datesArr.Length; i++)
                    {
                        qlDates.Add(QLHelper.ToQLDate(datesArr[i]));
                        qlFactors.Add(factorsArr[i]);
                    }

                    // DiscountCurve = InterpolatedDiscountCurve<LogLinear>.
                    var curve = new DiscountCurve(qlDates, qlFactors, dc, new TARGET());
                    curve.enableExtrapolation();
                    ObjectCache.StoreCurve(key, new YieldTermStructureHandle(curve));
                }

                return key;
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }
    }
}
