/*
 * SwapFunctions.cs — Vanilla interest-rate swap and OIS pricing.
 * Mirrors ql/instruments/vanillaswap.* and ql/pricingengines/swap/discountingswapengine.*
 */

using System;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.Swaps
{
    public static class SwapFunctions
    {
        // ─── Internal builder ─────────────────────────────────────────────────────

        /// <summary>
        /// Build a standard fixed-for-floating Euribor6M vanilla swap priced on a flat curve.
        /// Returns the configured swap object.
        /// </summary>
        private static VanillaSwap BuildSwap(
            double notional, double fixedRate,
            double riskFreeRate, Date evalQL, Date matQL,
            string fixedFrequency, string floatFrequency)
        {
            Settings.instance().setEvaluationDate(evalQL);
            var calendar   = new TARGET();
            var dayCounter = new Actual365Fixed();
            var discCurve  = QLHelper.FlatCurve(evalQL, riskFreeRate);

            var effectiveDate = calendar.advance(evalQL, 2, TimeUnit.Days);

            var fixedFreq = QLHelper.ParseFrequency(fixedFrequency);
            var floatFreq = QLHelper.ParseFrequency(floatFrequency);

            var fixedSchedule = new Schedule(effectiveDate, matQL,
                new Period(fixedFreq), calendar,
                BusinessDayConvention.ModifiedFollowing,
                BusinessDayConvention.ModifiedFollowing,
                DateGeneration.Rule.Forward, false);

            var floatSchedule = new Schedule(effectiveDate, matQL,
                new Period(floatFreq), calendar,
                BusinessDayConvention.ModifiedFollowing,
                BusinessDayConvention.ModifiedFollowing,
                DateGeneration.Rule.Forward, false);

            // Euribor6M index linked to the flat discount curve for forward rate estimation.
            var index = new Euribor6M(discCurve);

            var swap = new VanillaSwap(
                VanillaSwap.Type.Payer, notional,
                fixedSchedule, fixedRate, new Thirty360(Thirty360.Convention.BondBasis),
                floatSchedule, index, 0.0, index.dayCounter());

            swap.setPricingEngine(new DiscountingSwapEngine(discCurve));
            return swap;
        }

        // ─── UDFs ─────────────────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_SwapNPV",
                       Category = "QuantLib — Swaps",
                       Description = "NPV of a fixed-for-floating Euribor6M swap (flat discount curve).\n" +
                                     "Positive = Payer (pays fixed, receives floating) from buyer's perspective.")]
        public static object QL_SwapNPV(
            [ExcelArgument(Description = "Notional amount")] double notional,
            [ExcelArgument(Description = "Fixed coupon rate (decimal, e.g. 0.03)")] double fixedRate,
            [ExcelArgument(Description = "Flat risk-free / discount rate (decimal)")] double riskFreeRate,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Maturity date (Excel date)")] double maturityDate,
            [ExcelArgument(Description = "Fixed leg frequency (default Annual)")] string fixedFrequency = "Annual",
            [ExcelArgument(Description = "Floating leg frequency (default Semiannual)")] string floatFrequency = "Semiannual")
        {
            try
            {
                var evalQL = QLHelper.ToQLDate(evalDate);
                var matQL  = QLHelper.ToQLDate(maturityDate);
                var swap   = BuildSwap(notional, fixedRate, riskFreeRate, evalQL, matQL,
                                       fixedFrequency, floatFrequency);
                return swap.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_SwapFairRate",
                       Category = "QuantLib — Swaps",
                       Description = "Fair fixed rate of a vanilla swap (flat curve). This is the fixed rate that gives zero NPV.")]
        public static object QL_SwapFairRate(
            [ExcelArgument(Description = "Notional amount")] double notional,
            [ExcelArgument(Description = "Flat discount rate (decimal)")] double riskFreeRate,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Maturity date (Excel date)")] double maturityDate,
            [ExcelArgument(Description = "Fixed leg frequency (default Annual)")] string fixedFrequency = "Annual",
            [ExcelArgument(Description = "Floating leg frequency (default Semiannual)")] string floatFrequency = "Semiannual")
        {
            try
            {
                var evalQL = QLHelper.ToQLDate(evalDate);
                var matQL  = QLHelper.ToQLDate(maturityDate);
                // Pass 0 as fixed rate — fairRate() recalculates regardless.
                var swap = BuildSwap(notional, 0.0, riskFreeRate, evalQL, matQL,
                                     fixedFrequency, floatFrequency);
                return swap.fairRate();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_SwapFairSpread",
                       Category = "QuantLib — Swaps",
                       Description = "Fair floating spread (over the index) that makes the swap NPV zero.")]
        public static object QL_SwapFairSpread(
            [ExcelArgument(Description = "Notional amount")] double notional,
            [ExcelArgument(Description = "Fixed coupon rate (decimal)")] double fixedRate,
            [ExcelArgument(Description = "Flat discount rate (decimal)")] double riskFreeRate,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Maturity date (Excel date)")] double maturityDate,
            [ExcelArgument(Description = "Fixed leg frequency (default Annual)")] string fixedFrequency = "Annual",
            [ExcelArgument(Description = "Floating leg frequency (default Semiannual)")] string floatFrequency = "Semiannual")
        {
            try
            {
                var evalQL = QLHelper.ToQLDate(evalDate);
                var matQL  = QLHelper.ToQLDate(maturityDate);
                var swap   = BuildSwap(notional, fixedRate, riskFreeRate, evalQL, matQL,
                                       fixedFrequency, floatFrequency);
                return swap.fairSpread();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_SwapFixedLegNPV",
                       Category = "QuantLib — Swaps",
                       Description = "NPV of the fixed leg of a vanilla swap (flat curve).")]
        public static object QL_SwapFixedLegNPV(
            [ExcelArgument(Description = "Notional amount")] double notional,
            [ExcelArgument(Description = "Fixed coupon rate (decimal)")] double fixedRate,
            [ExcelArgument(Description = "Flat discount rate (decimal)")] double riskFreeRate,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Fixed frequency (default Annual)")] string fixedFrequency = "Annual",
            [ExcelArgument(Description = "Float frequency (default Semiannual)")] string floatFrequency = "Semiannual")
        {
            try
            {
                var swap = BuildSwap(notional, fixedRate, riskFreeRate,
                                     QLHelper.ToQLDate(evalDate), QLHelper.ToQLDate(maturityDate),
                                     fixedFrequency, floatFrequency);
                return swap.fixedLegNPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_SwapFloatLegNPV",
                       Category = "QuantLib — Swaps",
                       Description = "NPV of the floating leg of a vanilla swap (flat curve).")]
        public static object QL_SwapFloatLegNPV(
            [ExcelArgument(Description = "Notional amount")] double notional,
            [ExcelArgument(Description = "Fixed coupon rate (decimal)")] double fixedRate,
            [ExcelArgument(Description = "Flat discount rate (decimal)")] double riskFreeRate,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Fixed frequency (default Annual)")] string fixedFrequency = "Annual",
            [ExcelArgument(Description = "Float frequency (default Semiannual)")] string floatFrequency = "Semiannual")
        {
            try
            {
                var swap = BuildSwap(notional, fixedRate, riskFreeRate,
                                     QLHelper.ToQLDate(evalDate), QLHelper.ToQLDate(maturityDate),
                                     fixedFrequency, floatFrequency);
                return swap.floatingLegNPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_SwapFixedLegBPS",
                       Category = "QuantLib — Swaps",
                       Description = "BPS (DV01) of the fixed leg — NPV change for a 1bp shift in fixed rate.")]
        public static object QL_SwapFixedLegBPS(
            [ExcelArgument(Description = "Notional amount")] double notional,
            [ExcelArgument(Description = "Fixed coupon rate (decimal)")] double fixedRate,
            [ExcelArgument(Description = "Flat discount rate (decimal)")] double riskFreeRate,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Fixed frequency (default Annual)")] string fixedFrequency = "Annual",
            [ExcelArgument(Description = "Float frequency (default Semiannual)")] string floatFrequency = "Semiannual")
        {
            try
            {
                var swap = BuildSwap(notional, fixedRate, riskFreeRate,
                                     QLHelper.ToQLDate(evalDate), QLHelper.ToQLDate(maturityDate),
                                     fixedFrequency, floatFrequency);
                return swap.fixedLegBPS();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }
    }
}
