/*
 * ShortRateFunctions.cs — Short-rate model pricing (Hull-White, Vasicek, G2).
 * Mirrors ql/models/shortrate/ and ql/pricingengines/swaption/jamshidianswaptionengine.*
 */

using System;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.Models
{
    public static class ShortRateFunctions
    {
        // ─── Hull-White ───────────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_HullWhiteDiscountBond",
                       Category = "QuantLib — Short-Rate Models",
                       Description = "Discount bond price P(T1,T2) from the Hull-White model with a flat initial curve.\n" +
                                     "HW: dr = (θ(t) − a·r) dt + σ dW\n" +
                                     "a     = mean-reversion speed\n" +
                                     "sigma = instantaneous short-rate volatility")]
        public static object QL_HullWhiteDiscountBond(
            [ExcelArgument(Description = "Flat initial short rate (decimal)")] double initialRate,
            [ExcelArgument(Description = "Mean-reversion speed a")] double a,
            [ExcelArgument(Description = "Short-rate volatility σ")] double sigma,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Bond start date T1")] double bondStart,
            [ExcelArgument(Description = "Bond maturity date T2")] double bondMaturity)
        {
            try
            {
                var evalQL = QLHelper.ToQLDate(evalDate);
                var curve  = QLHelper.FlatCurve(evalQL, initialRate);
                var hw     = new HullWhite(curve, a, sigma);
                var t1     = new Actual365Fixed().yearFraction(evalQL, QLHelper.ToQLDate(bondStart));
                var t2     = new Actual365Fixed().yearFraction(evalQL, QLHelper.ToQLDate(bondMaturity));
                return hw.discountBond(t1, t2, initialRate);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_HullWhiteSwaption",
                       Category = "QuantLib — Short-Rate Models",
                       Description = "European swaption price under the Hull-White model (Jamshidian engine).\n" +
                                     "The underlying swap: Payer, notional=1, annual fixed vs Euribor6M.")]
        public static object QL_HullWhiteSwaption(
            [ExcelArgument(Description = "Flat initial short rate (decimal)")] double initialRate,
            [ExcelArgument(Description = "Mean-reversion speed a")] double a,
            [ExcelArgument(Description = "Short-rate volatility σ")] double sigma,
            [ExcelArgument(Description = "Fixed coupon rate of underlying swap")] double fixedRate,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Swaption expiry date (option exercise)")] double expiryDate,
            [ExcelArgument(Description = "Swap maturity date")] double swapMaturity)
        {
            try
            {
                var evalQL   = QLHelper.ToQLDate(evalDate);
                var expiryQL = QLHelper.ToQLDate(expiryDate);
                var matQL    = QLHelper.ToQLDate(swapMaturity);
                Settings.instance().setEvaluationDate(evalQL);

                var discCurve = QLHelper.FlatCurve(evalQL, initialRate);
                var hw        = new HullWhite(discCurve, a, sigma);
                var calendar  = new TARGET();
                var dayCounter = new Actual365Fixed();

                var effectiveDate = calendar.advance(evalQL, 2, TimeUnit.Days);
                var floatSchedule = new Schedule(effectiveDate, matQL, new Period(Frequency.Semiannual),
                    calendar, BusinessDayConvention.ModifiedFollowing,
                    BusinessDayConvention.ModifiedFollowing,
                    DateGeneration.Rule.Forward, false);
                var fixedSchedule = new Schedule(effectiveDate, matQL, new Period(Frequency.Annual),
                    calendar, BusinessDayConvention.ModifiedFollowing,
                    BusinessDayConvention.ModifiedFollowing,
                    DateGeneration.Rule.Forward, false);

                var index  = new Euribor6M(discCurve);
                var swap   = new VanillaSwap(
                    VanillaSwap.Type.Payer, 1.0,
                    fixedSchedule, fixedRate, new Thirty360(Thirty360.Convention.BondBasis),
                    floatSchedule, index, 0.0, index.dayCounter());

                var swaption = new Swaption(swap, new EuropeanExercise(expiryQL));
                swaption.setPricingEngine(new JamshidianSwaptionEngine(hw));
                return swaption.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Vasicek ──────────────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_VasicekDiscountBond",
                       Category = "QuantLib — Short-Rate Models",
                       Description = "Vasicek discount bond P(T1,T2) given current short rate r0.\n" +
                                     "Vasicek: dr = a(b − r) dt + σ dW\n" +
                                     "a = mean-reversion speed, b = long-run rate, sigma = short-rate vol.")]
        public static object QL_VasicekDiscountBond(
            [ExcelArgument(Description = "Current short rate r₀")] double r0,
            [ExcelArgument(Description = "Mean-reversion speed a")] double a,
            [ExcelArgument(Description = "Long-run rate b")] double b,
            [ExcelArgument(Description = "Short-rate volatility σ")] double sigma,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Bond start date T1")] double bondStart,
            [ExcelArgument(Description = "Bond maturity date T2")] double bondMaturity,
            [ExcelArgument(Description = "Risk premium λ (default 0)")] double lambda = 0.0)
        {
            try
            {
                var evalQL = QLHelper.ToQLDate(evalDate);
                var curve  = QLHelper.FlatCurve(evalQL, r0);
                var vasicek = new Vasicek(r0, a, b, sigma, lambda);
                var t1      = new Actual365Fixed().yearFraction(evalQL, QLHelper.ToQLDate(bondStart));
                var t2      = new Actual365Fixed().yearFraction(evalQL, QLHelper.ToQLDate(bondMaturity));
                return vasicek.discountBond(t1, t2, r0);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── G2 (two-factor Gaussian) ─────────────────────────────────────────────

        [ExcelFunction(Name = "QL_G2Swaption",
                       Category = "QuantLib — Short-Rate Models",
                       Description = "European swaption price under the two-factor Gaussian G2++ model.\n" +
                                     "G2: r(t) = x(t) + y(t) + φ(t)\n" +
                                     "dx = −a x dt + σ₁ dW₁\n" +
                                     "dy = −b y dt + σ₂ dW₂   corr = ρ\n" +
                                     "a,b = mean reversion; σ1,σ2 = factor vols; rho = correlation.")]
        public static object QL_G2Swaption(
            [ExcelArgument(Description = "Flat initial rate")] double initialRate,
            [ExcelArgument(Description = "First factor mean-reversion a")] double a,
            [ExcelArgument(Description = "First factor vol σ₁")] double sigma1,
            [ExcelArgument(Description = "Second factor mean-reversion b")] double b,
            [ExcelArgument(Description = "Second factor vol σ₂")] double sigma2,
            [ExcelArgument(Description = "Factor correlation ρ")] double rho,
            [ExcelArgument(Description = "Fixed coupon rate of underlying swap")] double fixedRate,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Swaption expiry date")] double expiryDate,
            [ExcelArgument(Description = "Swap maturity date")] double swapMaturity)
        {
            try
            {
                var evalQL   = QLHelper.ToQLDate(evalDate);
                var expiryQL = QLHelper.ToQLDate(expiryDate);
                var matQL    = QLHelper.ToQLDate(swapMaturity);
                Settings.instance().setEvaluationDate(evalQL);

                var discCurve = QLHelper.FlatCurve(evalQL, initialRate);
                var g2        = new G2(discCurve, a, sigma1, b, sigma2, rho);
                var calendar  = new TARGET();

                var effectiveDate = calendar.advance(evalQL, 2, TimeUnit.Days);
                var fixedSchedule = new Schedule(effectiveDate, matQL, new Period(Frequency.Annual),
                    calendar, BusinessDayConvention.ModifiedFollowing,
                    BusinessDayConvention.ModifiedFollowing,
                    DateGeneration.Rule.Forward, false);
                var floatSchedule = new Schedule(effectiveDate, matQL, new Period(Frequency.Semiannual),
                    calendar, BusinessDayConvention.ModifiedFollowing,
                    BusinessDayConvention.ModifiedFollowing,
                    DateGeneration.Rule.Forward, false);

                var index  = new Euribor6M(discCurve);
                var swap   = new VanillaSwap(
                    VanillaSwap.Type.Payer, 1.0,
                    fixedSchedule, fixedRate, new Thirty360(Thirty360.Convention.BondBasis),
                    floatSchedule, index, 0.0, index.dayCounter());

                var swaption = new Swaption(swap, new EuropeanExercise(expiryQL));
                swaption.setPricingEngine(new G2SwaptionEngine(g2, 6, 16));
                return swaption.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Black swaption (market convention) ──────────────────────────────────

        [ExcelFunction(Name = "QL_BlackSwaption",
                       Category = "QuantLib — Short-Rate Models",
                       Description = "European swaption price using the Black (log-normal) model.\n" +
                                     "This is the standard market-convention swaption pricing formula.")]
        public static object QL_BlackSwaption(
            [ExcelArgument(Description = "Flat discount rate")] double discountRate,
            [ExcelArgument(Description = "Fixed coupon rate of underlying swap")] double fixedRate,
            [ExcelArgument(Description = "Black volatility for the swaption")] double swaptionVol,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Swaption expiry date")] double expiryDate,
            [ExcelArgument(Description = "Swap maturity date")] double swapMaturity)
        {
            try
            {
                var evalQL   = QLHelper.ToQLDate(evalDate);
                var expiryQL = QLHelper.ToQLDate(expiryDate);
                var matQL    = QLHelper.ToQLDate(swapMaturity);
                Settings.instance().setEvaluationDate(evalQL);

                var discCurve = QLHelper.FlatCurve(evalQL, discountRate);
                var calendar  = new TARGET();
                var dayCounter = new Actual365Fixed();

                var effectiveDate = calendar.advance(evalQL, 2, TimeUnit.Days);
                var fixedSchedule = new Schedule(effectiveDate, matQL, new Period(Frequency.Annual),
                    calendar, BusinessDayConvention.ModifiedFollowing,
                    BusinessDayConvention.ModifiedFollowing,
                    DateGeneration.Rule.Forward, false);
                var floatSchedule = new Schedule(effectiveDate, matQL, new Period(Frequency.Semiannual),
                    calendar, BusinessDayConvention.ModifiedFollowing,
                    BusinessDayConvention.ModifiedFollowing,
                    DateGeneration.Rule.Forward, false);

                var index  = new Euribor6M(discCurve);
                var swap   = new VanillaSwap(
                    VanillaSwap.Type.Payer, 1.0,
                    fixedSchedule, fixedRate, new Thirty360(Thirty360.Convention.BondBasis),
                    floatSchedule, index, 0.0, index.dayCounter());

                var volHandle  = new SwaptionVolatilityStructureHandle(
                    new ConstantSwaptionVolatility(evalQL, new NullCalendar(),
                        BusinessDayConvention.Unadjusted, swaptionVol, dayCounter));

                var swaption = new Swaption(swap, new EuropeanExercise(expiryQL));
                swaption.setPricingEngine(new BlackSwaptionEngine(discCurve, volHandle));
                return swaption.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }
    }
}
