/*
 * BondFunctions.cs — Bond pricing, yield, duration, convexity, and BPS.
 * Mirrors ql/instruments/bond.*, ql/instruments/fixedratebond.*,
 *         ql/instruments/zerocouponbond.*, ql/pricingengines/bond/*.
 *
 * API notes (from SWIG wrapper inspection):
 *   - FixedRateBond/ZeroCouponBond take int (not uint) for settlementDays
 *   - BondFunctions.yield()  requires a BondPrice wrapper (not a raw double)
 *   - BondFunctions.accruedAmount() is the correct method name (not accruedInterest)
 *   - BondFunctions.bps() accepts a YieldTermStructure (use handle.currentLink())
 */

using System;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.Bonds
{
    public static class BondFunctions
    {
        // ─── Schedule builder ─────────────────────────────────────────────────────

        private static Schedule MakeSchedule(Date effectiveDate, Date maturityDate,
                                             Frequency frequency, Calendar calendar)
        {
            return new Schedule(effectiveDate, maturityDate, new Period(frequency), calendar,
                                BusinessDayConvention.Unadjusted, BusinessDayConvention.Unadjusted,
                                DateGeneration.Rule.Backward, false);
        }

        // ─── Zero coupon bond ─────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_ZeroCouponBondPrice",
                       Category = "QuantLib — Bonds",
                       Description = "Dirty price (NPV) of a zero-coupon bond discounted by a flat yield curve.")]
        public static object QL_ZeroCouponBondPrice(
            [ExcelArgument(Description = "Face amount (e.g. 100)")] double faceAmount,
            [ExcelArgument(Description = "Issue date (Excel date)")] double issueDate,
            [ExcelArgument(Description = "Maturity date (Excel date)")] double maturityDate,
            [ExcelArgument(Description = "Flat discount rate (decimal)")] double discountRate,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Settlement days (default 3)")] int settlementDays = 3)
        {
            try
            {
                var evalQL    = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var discCurve = QLHelper.FlatCurve(evalQL, discountRate);
                var bond      = new ZeroCouponBond(
                    (uint)settlementDays, new TARGET(), faceAmount,
                    QLHelper.ToQLDate(maturityDate),
                    BusinessDayConvention.Following, 100.0, QLHelper.ToQLDate(issueDate));
                bond.setPricingEngine(new DiscountingBondEngine(discCurve));
                return bond.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_ZeroCouponBondYield",
                       Category = "QuantLib — Bonds",
                       Description = "Yield to maturity of a zero-coupon bond from its clean price.")]
        public static object QL_ZeroCouponBondYield(
            [ExcelArgument(Description = "Face amount")] double faceAmount,
            [ExcelArgument(Description = "Issue date (Excel date)")] double issueDate,
            [ExcelArgument(Description = "Maturity date (Excel date)")] double maturityDate,
            [ExcelArgument(Description = "Clean price")] double cleanPrice,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Compounding convention (default Compounded)")] string compounding = "Compounded",
            [ExcelArgument(Description = "Frequency (default Annual)")] string frequency = "Annual",
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365",
            [ExcelArgument(Description = "Settlement days (default 3)")] int settlementDays = 3)
        {
            try
            {
                var evalQL    = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var bond      = new ZeroCouponBond(
                    (uint)settlementDays, new TARGET(), faceAmount,
                    QLHelper.ToQLDate(maturityDate),
                    BusinessDayConvention.Following, 100.0, QLHelper.ToQLDate(issueDate));
                var dc        = QLHelper.ParseDayCounter(dayCounter);
                var comp      = QLHelper.ParseCompounding(compounding);
                var freq      = QLHelper.ParseFrequency(frequency);
                var settlDate = new TARGET().advance(evalQL, settlementDays, TimeUnit.Days);
                return QuantLib.BondFunctions.yield(bond,
                    new BondPrice(cleanPrice, BondPrice.Type.Clean), dc, comp, freq, settlDate);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Fixed-rate bond ──────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_FixedRateBondPrice",
                       Category = "QuantLib — Bonds",
                       Description = "Clean price of a fixed-rate coupon bond discounted by a flat yield curve.")]
        public static object QL_FixedRateBondPrice(
            [ExcelArgument(Description = "Face amount (e.g. 100)")] double faceAmount,
            [ExcelArgument(Description = "Annual coupon rate (decimal, e.g. 0.05)")] double couponRate,
            [ExcelArgument(Description = "Issue date (Excel date)")] double issueDate,
            [ExcelArgument(Description = "Maturity date (Excel date)")] double maturityDate,
            [ExcelArgument(Description = "Flat discount rate (decimal)")] double discountRate,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Coupon frequency (default Semiannual)")] string frequency = "Semiannual",
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365",
            [ExcelArgument(Description = "Settlement days (default 3)")] int settlementDays = 3)
        {
            try
            {
                var evalQL    = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var discCurve = QLHelper.FlatCurve(evalQL, discountRate);
                var freq      = QLHelper.ParseFrequency(frequency);
                var dc        = QLHelper.ParseDayCounter(dayCounter);
                var schedule  = MakeSchedule(QLHelper.ToQLDate(issueDate), QLHelper.ToQLDate(maturityDate),
                                             freq, new TARGET());
                var bond = new FixedRateBond(settlementDays, faceAmount, schedule,
                               QLHelper.ToDoubleVector(new[] { couponRate }), dc);
                bond.setPricingEngine(new DiscountingBondEngine(discCurve));
                return bond.cleanPrice();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_FixedRateBondNPV",
                       Category = "QuantLib — Bonds",
                       Description = "Dirty (full) price (NPV) of a fixed-rate bond discounted by a flat yield curve.")]
        public static object QL_FixedRateBondNPV(
            [ExcelArgument(Description = "Face amount")] double faceAmount,
            [ExcelArgument(Description = "Annual coupon rate (decimal)")] double couponRate,
            [ExcelArgument(Description = "Issue date")] double issueDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Flat discount rate")] double discountRate,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Coupon frequency (default Semiannual)")] string frequency = "Semiannual",
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365",
            [ExcelArgument(Description = "Settlement days (default 3)")] int settlementDays = 3)
        {
            try
            {
                var evalQL    = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var discCurve = QLHelper.FlatCurve(evalQL, discountRate);
                var freq      = QLHelper.ParseFrequency(frequency);
                var dc        = QLHelper.ParseDayCounter(dayCounter);
                var schedule  = MakeSchedule(QLHelper.ToQLDate(issueDate), QLHelper.ToQLDate(maturityDate),
                                             freq, new TARGET());
                var bond = new FixedRateBond(settlementDays, faceAmount, schedule,
                               QLHelper.ToDoubleVector(new[] { couponRate }), dc);
                bond.setPricingEngine(new DiscountingBondEngine(discCurve));
                return bond.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_FixedRateBondYield",
                       Category = "QuantLib — Bonds",
                       Description = "Yield to maturity of a fixed-rate bond from its clean price.")]
        public static object QL_FixedRateBondYield(
            [ExcelArgument(Description = "Face amount")] double faceAmount,
            [ExcelArgument(Description = "Annual coupon rate (decimal)")] double couponRate,
            [ExcelArgument(Description = "Issue date")] double issueDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Clean price")] double cleanPrice,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Coupon frequency (default Semiannual)")] string frequency = "Semiannual",
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365",
            [ExcelArgument(Description = "Compounding convention (default Compounded)")] string compounding = "Compounded",
            [ExcelArgument(Description = "Settlement days (default 3)")] int settlementDays = 3)
        {
            try
            {
                var evalQL   = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var freq     = QLHelper.ParseFrequency(frequency);
                var dc       = QLHelper.ParseDayCounter(dayCounter);
                var comp     = QLHelper.ParseCompounding(compounding);
                var schedule = MakeSchedule(QLHelper.ToQLDate(issueDate), QLHelper.ToQLDate(maturityDate),
                                            freq, new TARGET());
                var bond     = new FixedRateBond(settlementDays, faceAmount, schedule,
                                   QLHelper.ToDoubleVector(new[] { couponRate }), dc);
                var settlDate = new TARGET().advance(evalQL, settlementDays, TimeUnit.Days);
                return QuantLib.BondFunctions.yield(bond,
                    new BondPrice(cleanPrice, BondPrice.Type.Clean), dc, comp, freq, settlDate);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_FixedRateBondDuration",
                       Category = "QuantLib — Bonds",
                       Description = "Duration of a fixed-rate bond.\n" +
                                     "DurationType: \"Modified\" (default), \"Macaulay\", or \"Simple\".")]
        public static object QL_FixedRateBondDuration(
            [ExcelArgument(Description = "Face amount")] double faceAmount,
            [ExcelArgument(Description = "Annual coupon rate (decimal)")] double couponRate,
            [ExcelArgument(Description = "Issue date")] double issueDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Yield to maturity (decimal)")] double ytm,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Duration type: \"Modified\", \"Macaulay\", or \"Simple\" (default Modified)")] string durationType = "Modified",
            [ExcelArgument(Description = "Coupon frequency (default Semiannual)")] string frequency = "Semiannual",
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365",
            [ExcelArgument(Description = "Compounding convention (default Compounded)")] string compounding = "Compounded",
            [ExcelArgument(Description = "Settlement days (default 3)")] int settlementDays = 3)
        {
            try
            {
                var evalQL   = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var freq     = QLHelper.ParseFrequency(frequency);
                var dc       = QLHelper.ParseDayCounter(dayCounter);
                var comp     = QLHelper.ParseCompounding(compounding);
                var schedule = MakeSchedule(QLHelper.ToQLDate(issueDate), QLHelper.ToQLDate(maturityDate),
                                            freq, new TARGET());
                var bond     = new FixedRateBond(settlementDays, faceAmount, schedule,
                                   QLHelper.ToDoubleVector(new[] { couponRate }), dc);
                var ir       = new InterestRate(ytm, dc, comp, freq);
                var dt       = durationType.Trim().ToLowerInvariant() switch
                {
                    "macaulay" => Duration.Type.Macaulay,
                    "simple"   => Duration.Type.Simple,
                    _          => Duration.Type.Modified,
                };
                var settlDate = new TARGET().advance(evalQL, settlementDays, TimeUnit.Days);
                return QuantLib.BondFunctions.duration(bond, ir, dt, settlDate);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_FixedRateBondConvexity",
                       Category = "QuantLib — Bonds",
                       Description = "Convexity of a fixed-rate bond (d²P/dY² / P).")]
        public static object QL_FixedRateBondConvexity(
            [ExcelArgument(Description = "Face amount")] double faceAmount,
            [ExcelArgument(Description = "Annual coupon rate (decimal)")] double couponRate,
            [ExcelArgument(Description = "Issue date")] double issueDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Yield to maturity (decimal)")] double ytm,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Coupon frequency (default Semiannual)")] string frequency = "Semiannual",
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365",
            [ExcelArgument(Description = "Compounding convention (default Compounded)")] string compounding = "Compounded",
            [ExcelArgument(Description = "Settlement days (default 3)")] int settlementDays = 3)
        {
            try
            {
                var evalQL   = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var freq     = QLHelper.ParseFrequency(frequency);
                var dc       = QLHelper.ParseDayCounter(dayCounter);
                var comp     = QLHelper.ParseCompounding(compounding);
                var schedule = MakeSchedule(QLHelper.ToQLDate(issueDate), QLHelper.ToQLDate(maturityDate),
                                            freq, new TARGET());
                var bond     = new FixedRateBond(settlementDays, faceAmount, schedule,
                                   QLHelper.ToDoubleVector(new[] { couponRate }), dc);
                var ir       = new InterestRate(ytm, dc, comp, freq);
                var settlDate = new TARGET().advance(evalQL, settlementDays, TimeUnit.Days);
                return QuantLib.BondFunctions.convexity(bond, ir, settlDate);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_FixedRateBondBPS",
                       Category = "QuantLib — Bonds",
                       Description = "Basis-point sensitivity (DV01) of a fixed-rate bond.")]
        public static object QL_FixedRateBondBPS(
            [ExcelArgument(Description = "Face amount")] double faceAmount,
            [ExcelArgument(Description = "Annual coupon rate (decimal)")] double couponRate,
            [ExcelArgument(Description = "Issue date")] double issueDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Flat discount rate")] double discountRate,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Coupon frequency (default Semiannual)")] string frequency = "Semiannual",
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365",
            [ExcelArgument(Description = "Settlement days (default 3)")] int settlementDays = 3)
        {
            try
            {
                var evalQL    = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var discCurve = QLHelper.FlatCurve(evalQL, discountRate);
                var freq      = QLHelper.ParseFrequency(frequency);
                var dc        = QLHelper.ParseDayCounter(dayCounter);
                var schedule  = MakeSchedule(QLHelper.ToQLDate(issueDate), QLHelper.ToQLDate(maturityDate),
                                             freq, new TARGET());
                var bond      = new FixedRateBond(settlementDays, faceAmount, schedule,
                                    QLHelper.ToDoubleVector(new[] { couponRate }), dc);
                bond.setPricingEngine(new DiscountingBondEngine(discCurve));
                var settlDate = new TARGET().advance(evalQL, settlementDays, TimeUnit.Days);
                return QuantLib.BondFunctions.bps(bond, discCurve.currentLink(), settlDate);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BondAccruedAmount",
                       Category = "QuantLib — Bonds",
                       Description = "Accrued coupon amount of a fixed-rate bond at the settlement date.")]
        public static object QL_BondAccruedAmount(
            [ExcelArgument(Description = "Face amount")] double faceAmount,
            [ExcelArgument(Description = "Annual coupon rate (decimal)")] double couponRate,
            [ExcelArgument(Description = "Issue date")] double issueDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Coupon frequency (default Semiannual)")] string frequency = "Semiannual",
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365",
            [ExcelArgument(Description = "Settlement days (default 3)")] int settlementDays = 3)
        {
            try
            {
                var evalQL   = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var freq     = QLHelper.ParseFrequency(frequency);
                var dc       = QLHelper.ParseDayCounter(dayCounter);
                var schedule = MakeSchedule(QLHelper.ToQLDate(issueDate), QLHelper.ToQLDate(maturityDate),
                                            freq, new TARGET());
                var bond     = new FixedRateBond(settlementDays, faceAmount, schedule,
                                   QLHelper.ToDoubleVector(new[] { couponRate }), dc);
                var settlDate = new TARGET().advance(evalQL, settlementDays, TimeUnit.Days);
                return QuantLib.BondFunctions.accruedAmount(bond, settlDate);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }
    }
}
