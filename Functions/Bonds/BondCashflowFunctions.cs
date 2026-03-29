/*
 * BondCashflowFunctions.cs — Bond and FRN pricing with full cashflow schedules.
 * Returns price (clean / dirty / NPV) and per-cashflow tables of
 * [Payment Date, Amount, Discount Factor, Present Value].
 *
 * Curve inputs:
 *   - Flat discount rate: single number used for discounting and (for FRN) forward estimation.
 *   - Bootstrapped swap curve: pass tenor/rate arrays from QL_SwapCurve* functions indirectly
 *     via the flat rate approximation for now.
 */

using System;
using SM = System.Math;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.Bonds
{
    public static class BondCashflowFunctions
    {
        // ─── Shared schedule builder ──────────────────────────────────────────────

        private static Schedule MakeSchedule(Date effectiveDate, Date maturityDate,
                                             Frequency frequency, Calendar calendar)
            => new Schedule(effectiveDate, maturityDate, new Period(frequency), calendar,
                            BusinessDayConvention.Unadjusted, BusinessDayConvention.Unadjusted,
                            DateGeneration.Rule.Backward, false);

        // ─── Fixed-rate bond ──────────────────────────────────────────────────────

        /// <summary>Build and price a FixedRateBond; returns (bond, discCurve).</summary>
        private static (FixedRateBond bond, YieldTermStructureHandle curve) BuildFixed(
            double faceAmount, double couponRate,
            double issueDateSerial, double maturitySerial,
            double discountRate, Date evalQL,
            Frequency freq, DayCounter dc, int settlementDays)
        {
            var discCurve = QLHelper.FlatCurve(evalQL, discountRate);
            var schedule  = MakeSchedule(QLHelper.ToQLDate(issueDateSerial),
                                         QLHelper.ToQLDate(maturitySerial), freq, new TARGET());
            var bond = new FixedRateBond(settlementDays, faceAmount, schedule,
                           QLHelper.ToDoubleVector(new[] { couponRate }), dc);
            bond.setPricingEngine(new DiscountingBondEngine(discCurve));
            return (bond, discCurve);
        }

        [ExcelFunction(Name = "QL_FixedBondPrice",
                       Category = "QuantLib — Bonds",
                       Description = "Clean price, dirty price and accrued interest of a fixed-rate bond.\n" +
                                     "Returns a 1×3 array: [CleanPrice, DirtyPrice, AccruedInterest].\n" +
                                     "Enter as array formula over 1 row × 3 columns.")]
        public static object QL_FixedBondPrice(
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
                var evalQL = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var freq   = QLHelper.ParseFrequency(frequency);
                var dc     = QLHelper.ParseDayCounter(dayCounter);
                var (bond, curve) = BuildFixed(faceAmount, couponRate, issueDate, maturityDate,
                                               discountRate, evalQL, freq, dc, settlementDays);
                var settlDate = new TARGET().advance(evalQL, settlementDays, TimeUnit.Days);
                return new object[,]
                {
                    {
                        bond.cleanPrice(),
                        bond.dirtyPrice(),
                        QuantLib.BondFunctions.accruedAmount(bond, settlDate)
                    }
                };
            }
            catch (Exception ex) { return new object[,] { { $"#QL_ERR: {ex.Message}" } }; }
        }

        [ExcelFunction(Name = "QL_FixedBondCashflows",
                       Category = "QuantLib — Bonds",
                       Description = "Cashflow schedule of a fixed-rate bond.\n" +
                                     "Returns an array with columns: [Payment Date, Amount, Discount Factor, Present Value].\n" +
                                     "Enter as array formula over N rows × 4 columns.")]
        public static object QL_FixedBondCashflows(
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
                var evalQL = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var freq   = QLHelper.ParseFrequency(frequency);
                var dc     = QLHelper.ParseDayCounter(dayCounter);
                var (bond, curve) = BuildFixed(faceAmount, couponRate, issueDate, maturityDate,
                                               discountRate, evalQL, freq, dc, settlementDays);

                return BuildCashflowTable(bond.cashflows(), curve);
            }
            catch (Exception ex) { return new object[,] { { $"#QL_ERR: {ex.Message}" } }; }
        }

        // ─── Floating-rate note ───────────────────────────────────────────────────

        /// <summary>Build and price a FloatingRateBond (Euribor6M, flat curve).</summary>
        private static (FloatingRateBond bond, YieldTermStructureHandle curve) BuildFRN(
            double faceAmount, double spread,
            double issueDateSerial, double maturitySerial,
            double discountRate, Date evalQL,
            Frequency freq, int settlementDays)
        {
            var discCurve  = QLHelper.FlatCurve(evalQL, discountRate);
            var index      = new Euribor6M(discCurve);
            var calendar   = new TARGET();
            var schedule   = new Schedule(
                                 QLHelper.ToQLDate(issueDateSerial),
                                 QLHelper.ToQLDate(maturitySerial),
                                 new Period(freq), calendar,
                                 BusinessDayConvention.ModifiedFollowing,
                                 BusinessDayConvention.ModifiedFollowing,
                                 DateGeneration.Rule.Backward, false);

            var gearings = new DoubleVector(); gearings.Add(1.0);
            var spreads  = new DoubleVector(); spreads.Add(spread);
            var caps     = new DoubleVector();
            var floors   = new DoubleVector();

            var bond = new FloatingRateBond(
                (uint)settlementDays, faceAmount, schedule, index,
                new Actual360(),
                BusinessDayConvention.ModifiedFollowing,
                index.fixingDays(),
                gearings, spreads, caps, floors,
                false,    // inArrears
                100.0,    // redemption
                QLHelper.ToQLDate(issueDateSerial));

            bond.setPricingEngine(new DiscountingBondEngine(discCurve));
            return (bond, discCurve);
        }

        [ExcelFunction(Name = "QL_FRNPrice",
                       Category = "QuantLib — Bonds",
                       Description = "Clean price, dirty price and accrued interest of a floating-rate note (Euribor6M, flat curve).\n" +
                                     "Returns a 1×3 array: [CleanPrice, DirtyPrice, AccruedInterest].\n" +
                                     "Enter as array formula over 1 row × 3 columns.")]
        public static object QL_FRNPrice(
            [ExcelArgument(Description = "Face amount (e.g. 100)")] double faceAmount,
            [ExcelArgument(Description = "Spread over index (decimal, e.g. 0.005 for +50bp)")] double spread,
            [ExcelArgument(Description = "Issue date (Excel date)")] double issueDate,
            [ExcelArgument(Description = "Maturity date (Excel date)")] double maturityDate,
            [ExcelArgument(Description = "Flat discount / forecast rate (decimal)")] double discountRate,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Coupon frequency (default Semiannual)")] string frequency = "Semiannual",
            [ExcelArgument(Description = "Settlement days (default 3)")] int settlementDays = 3)
        {
            try
            {
                var evalQL = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var freq   = QLHelper.ParseFrequency(frequency);
                var (bond, curve) = BuildFRN(faceAmount, spread, issueDate, maturityDate,
                                             discountRate, evalQL, freq, settlementDays);
                var settlDate = new TARGET().advance(evalQL, settlementDays, TimeUnit.Days);
                return new object[,]
                {
                    {
                        bond.cleanPrice(),
                        bond.dirtyPrice(),
                        QuantLib.BondFunctions.accruedAmount(bond, settlDate)
                    }
                };
            }
            catch (Exception ex) { return new object[,] { { $"#QL_ERR: {ex.Message}" } }; }
        }

        [ExcelFunction(Name = "QL_FRNCashflows",
                       Category = "QuantLib — Bonds",
                       Description = "Cashflow schedule of a floating-rate note (Euribor6M, flat curve).\n" +
                                     "Returns columns: [Payment Date, Amount, Discount Factor, Present Value].\n" +
                                     "Coupon amounts are estimated from the forward curve. Enter as array formula.")]
        public static object QL_FRNCashflows(
            [ExcelArgument(Description = "Face amount (e.g. 100)")] double faceAmount,
            [ExcelArgument(Description = "Spread over index (decimal, e.g. 0.005 for +50bp)")] double spread,
            [ExcelArgument(Description = "Issue date (Excel date)")] double issueDate,
            [ExcelArgument(Description = "Maturity date (Excel date)")] double maturityDate,
            [ExcelArgument(Description = "Flat discount / forecast rate (decimal)")] double discountRate,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Coupon frequency (default Semiannual)")] string frequency = "Semiannual",
            [ExcelArgument(Description = "Settlement days (default 3)")] int settlementDays = 3)
        {
            try
            {
                var evalQL = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var freq   = QLHelper.ParseFrequency(frequency);
                var (bond, curve) = BuildFRN(faceAmount, spread, issueDate, maturityDate,
                                             discountRate, evalQL, freq, settlementDays);

                return BuildCashflowTable(bond.cashflows(), curve);
            }
            catch (Exception ex) { return new object[,] { { $"#QL_ERR: {ex.Message}" } }; }
        }

        // ─── Shared cashflow table builder ────────────────────────────────────────

        /// <summary>
        /// Iterates a QuantLib Leg and returns a 2-D array:
        ///   Col 0: Payment date as Excel serial (format as Date in Excel)
        ///   Col 1: Cashflow amount
        ///   Col 2: Discount factor at payment date
        ///   Col 3: Present value (amount × discount factor)
        /// </summary>
        private static object[,] BuildCashflowTable(Leg leg, YieldTermStructureHandle curve)
        {
            int n      = (int)leg.Count;
            var result = new object[n, 4];
            for (int i = 0; i < n; i++)
            {
                var cf       = leg[i];
                var payDate  = cf.date();
                double amount  = cf.amount();
                double df      = curve.currentLink().discount(payDate);
                result[i, 0] = QLHelper.ToExcelDate(payDate);
                result[i, 1] = amount;
                result[i, 2] = df;
                result[i, 3] = amount * df;
            }
            return result;
        }
    }
}
