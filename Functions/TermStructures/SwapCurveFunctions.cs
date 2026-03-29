/*
 * SwapCurveFunctions.cs — Bootstrapped swap curve using market par swap rates.
 * Uses SwapRateHelper + PiecewiseFlatForward to build a full yield curve from
 * quoted par swap rates, then exposes zero rates, discount factors, and forward rates.
 */

using System;
using SM = System.Math;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.TermStructures
{
    public static class SwapCurveFunctions
    {
        // ─── Internal bootstrapper ────────────────────────────────────────────────

        /// <summary>
        /// Bootstrap a piecewise-flat-forward yield curve from par swap rates.
        /// tenorYears: integer years for each swap (e.g. 1, 2, 3, 5, 7, 10).
        /// parRates:   market par swap rates as decimals (e.g. 0.045 for 4.5%).
        /// Returns the bootstrapped curve wrapped in a handle.
        /// </summary>
        private static PiecewiseFlatForward BuildSwapCurve(
            double[] tenorYears, double[] parRates, Date evalQL)
        {
            if (tenorYears.Length != parRates.Length)
                throw new ArgumentException("tenorYears and parRates must have the same number of elements.");
            if (tenorYears.Length == 0)
                throw new ArgumentException("At least one tenor/rate pair is required.");

            Settings.instance().setEvaluationDate(evalQL);

            var calendar  = new TARGET();
            var dc        = new Actual365Fixed();
            var fixedDC   = new Thirty360(Thirty360.Convention.BondBasis);
            var index     = new Euribor6M();   // no curve — bootstrapper wires it internally

            var helpers = new RateHelperVector();
            for (int i = 0; i < tenorYears.Length; i++)
            {
                int years  = (int)SM.Round(tenorYears[i]);
                var quote  = new QuoteHandle(new SimpleQuote(parRates[i]));
                var tenor  = new Period(years, TimeUnit.Years);
                helpers.Add(new SwapRateHelper(
                    quote, tenor, calendar,
                    Frequency.Annual,
                    BusinessDayConvention.ModifiedFollowing,
                    fixedDC, index));
            }

            return new PiecewiseFlatForward(evalQL, helpers, dc);
        }

        // ─── Zero rate ────────────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_SwapCurveZeroRate",
                       Category = "QuantLib — Swap Curve",
                       Description = "Bootstrap a swap curve from par rates and return the zero rate at a query date.\n" +
                                     "tenorYears: range of integer tenor years (e.g. 1,2,3,5,7,10).\n" +
                                     "parRates: range of market par swap rates (e.g. 0.045).")]
        public static object QL_SwapCurveZeroRate(
            [ExcelArgument(Description = "Range of tenor years (integers, e.g. 1 2 3 5 7 10)")] object tenorYears,
            [ExcelArgument(Description = "Range of par swap rates (decimals, same order as tenors)")] object parRates,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Query date to read the zero rate at")] double queryDate,
            [ExcelArgument(Description = "Compounding (default Continuous)")] string compounding = "Continuous",
            [ExcelArgument(Description = "Frequency (default Annual)")] string frequency = "Annual")
        {
            try
            {
                var evalQL  = QLHelper.ToQLDate(evalDate);
                var queryQL = QLHelper.ToQLDate(queryDate);
                var tenors  = QLHelper.ToDoubleArray(tenorYears);
                var rates   = QLHelper.ToDoubleArray(parRates);
                var curve   = BuildSwapCurve(tenors, rates, evalQL);
                var dc      = new Actual365Fixed();
                var comp    = QLHelper.ParseCompounding(compounding);
                var freq    = QLHelper.ParseFrequency(frequency);
                return curve.zeroRate(queryQL, dc, comp, freq).rate();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Discount factor ──────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_SwapCurveDiscount",
                       Category = "QuantLib — Swap Curve",
                       Description = "Bootstrap a swap curve from par rates and return the discount factor P(0,T) at a query date.")]
        public static object QL_SwapCurveDiscount(
            [ExcelArgument(Description = "Range of tenor years (integers)")] object tenorYears,
            [ExcelArgument(Description = "Range of par swap rates (decimals)")] object parRates,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Query date")] double queryDate)
        {
            try
            {
                var evalQL  = QLHelper.ToQLDate(evalDate);
                var queryQL = QLHelper.ToQLDate(queryDate);
                var tenors  = QLHelper.ToDoubleArray(tenorYears);
                var rates   = QLHelper.ToDoubleArray(parRates);
                var curve   = BuildSwapCurve(tenors, rates, evalQL);
                return curve.discount(queryQL);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Forward rate ─────────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_SwapCurveForwardRate",
                       Category = "QuantLib — Swap Curve",
                       Description = "Bootstrap a swap curve and return the forward rate between two dates.")]
        public static object QL_SwapCurveForwardRate(
            [ExcelArgument(Description = "Range of tenor years (integers)")] object tenorYears,
            [ExcelArgument(Description = "Range of par swap rates (decimals)")] object parRates,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Forward period start date")] double date1,
            [ExcelArgument(Description = "Forward period end date")] double date2,
            [ExcelArgument(Description = "Compounding (default Continuous)")] string compounding = "Continuous",
            [ExcelArgument(Description = "Frequency (default Annual)")] string frequency = "Annual")
        {
            try
            {
                var evalQL = QLHelper.ToQLDate(evalDate);
                var d1     = QLHelper.ToQLDate(date1);
                var d2     = QLHelper.ToQLDate(date2);
                var tenors = QLHelper.ToDoubleArray(tenorYears);
                var rates  = QLHelper.ToDoubleArray(parRates);
                var curve  = BuildSwapCurve(tenors, rates, evalQL);
                var dc     = new Actual365Fixed();
                var comp   = QLHelper.ParseCompounding(compounding);
                var freq   = QLHelper.ParseFrequency(frequency);
                return curve.forwardRate(d1, d2, dc, comp, freq).rate();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Full curve table ─────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_SwapCurveTable",
                       Category = "QuantLib — Swap Curve",
                       Description = "Bootstrap a swap curve and return a table of [TenorYears, ZeroRate, DiscountFactor] " +
                                     "at each input tenor. Enter as an array formula (Ctrl+Shift+Enter) over N rows x 3 columns.")]
        public static object QL_SwapCurveTable(
            [ExcelArgument(Description = "Range of tenor years (integers)")] object tenorYears,
            [ExcelArgument(Description = "Range of par swap rates (decimals)")] object parRates,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Compounding for zero rates (default Continuous)")] string compounding = "Continuous",
            [ExcelArgument(Description = "Frequency (default Annual)")] string frequency = "Annual")
        {
            try
            {
                var evalQL  = QLHelper.ToQLDate(evalDate);
                var tenors  = QLHelper.ToDoubleArray(tenorYears);
                var rates   = QLHelper.ToDoubleArray(parRates);
                var curve   = BuildSwapCurve(tenors, rates, evalQL);
                var dc      = new Actual365Fixed();
                var comp    = QLHelper.ParseCompounding(compounding);
                var freq    = QLHelper.ParseFrequency(frequency);

                int n       = tenors.Length;
                var result  = new object[n, 3];
                for (int i = 0; i < n; i++)
                {
                    int years       = (int)SM.Round(tenors[i]);
                    var queryQL     = evalQL + new Period(years, TimeUnit.Years);
                    result[i, 0]    = tenors[i];
                    result[i, 1]    = curve.zeroRate(queryQL, dc, comp, freq).rate();
                    result[i, 2]    = curve.discount(queryQL);
                }
                return result;
            }
            catch (Exception ex) { return new object[,] { { $"#QL_ERR: {ex.Message}" } }; }
        }

        // ─── Par rate check (round-trip validation) ───────────────────────────────

        [ExcelFunction(Name = "QL_SwapCurveParRate",
                       Category = "QuantLib — Swap Curve",
                       Description = "Bootstrap a swap curve and return the implied par swap rate at a given tenor (round-trip check). " +
                                     "Should match the input par rate for tenors used in calibration.")]
        public static object QL_SwapCurveParRate(
            [ExcelArgument(Description = "Range of tenor years (integers)")] object tenorYears,
            [ExcelArgument(Description = "Range of par swap rates (decimals)")] object parRates,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Tenor in years to query (integer)")] double queryTenorYears,
            [ExcelArgument(Description = "Fixed leg frequency (default Annual)")] string fixedFrequency = "Annual",
            [ExcelArgument(Description = "Floating leg frequency (default Semiannual)")] string floatFrequency = "Semiannual")
        {
            try
            {
                var evalQL   = QLHelper.ToQLDate(evalDate);
                var tenors   = QLHelper.ToDoubleArray(tenorYears);
                var rates    = QLHelper.ToDoubleArray(parRates);
                var curve    = BuildSwapCurve(tenors, rates, evalQL);
                var handle   = new YieldTermStructureHandle(curve);

                Settings.instance().setEvaluationDate(evalQL);
                var calendar = new TARGET();
                var matQL    = evalQL + new Period((int)SM.Round(queryTenorYears), TimeUnit.Years);
                var effDate  = calendar.advance(evalQL, 2, TimeUnit.Days);
                var fixedDC  = new Thirty360(Thirty360.Convention.BondBasis);
                var fixedFreq = QLHelper.ParseFrequency(fixedFrequency);
                var floatFreq = QLHelper.ParseFrequency(floatFrequency);

                var fixedSchedule = new Schedule(effDate, matQL,
                    new Period(fixedFreq), calendar,
                    BusinessDayConvention.ModifiedFollowing,
                    BusinessDayConvention.ModifiedFollowing,
                    DateGeneration.Rule.Forward, false);

                var floatSchedule = new Schedule(effDate, matQL,
                    new Period(floatFreq), calendar,
                    BusinessDayConvention.ModifiedFollowing,
                    BusinessDayConvention.ModifiedFollowing,
                    DateGeneration.Rule.Forward, false);

                var index = new Euribor6M(handle);
                var swap  = new VanillaSwap(
                    VanillaSwap.Type.Payer, 1_000_000.0,
                    fixedSchedule, 0.0, fixedDC,
                    floatSchedule, index, 0.0, index.dayCounter());
                swap.setPricingEngine(new DiscountingSwapEngine(handle));
                return swap.fairRate();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }
    }
}
