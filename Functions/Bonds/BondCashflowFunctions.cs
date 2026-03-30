/*
 * BondCashflowFunctions.cs — Bond and FRN pricing with full cashflow schedules.
 *
 * The "curve" parameter in every public function accepts either:
 *   - A curve handle string returned by QL_BuildSwapCurve / QL_BuildFlatCurve, OR
 *   - A plain decimal number (e.g. 0.045) which is treated as a flat discount rate.
 *
 * This means the same function works for quick one-off pricing (pass a number) and
 * efficient portfolio pricing (pass the cached handle once, reuse across all bonds).
 */

using System;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.Bonds
{
    public static class BondCashflowFunctions
    {
        // ─── Curve resolver ───────────────────────────────────────────────────────

        /// <summary>
        /// Accept either a cached curve handle (string) or a flat rate (double).
        /// Called at the top of every public function — never throws; errors propagate
        /// as exceptions that the caller's try/catch will turn into #QL_ERR strings.
        /// </summary>
        private static YieldTermStructureHandle ResolveCurve(object curveOrRate, Date evalQL)
        {
            if (curveOrRate is string handle && handle.Length > 0)
                return ObjectCache.GetCurve(handle);

            // ExcelDNA passes numbers as double; also handle boxed double from object[,]
            double rate = Convert.ToDouble(curveOrRate);
            return QLHelper.FlatCurve(evalQL, rate);
        }

        // ─── Shared helpers ───────────────────────────────────────────────────────

        private static Schedule MakeSchedule(Date effectiveDate, Date maturityDate,
                                             Frequency frequency, Calendar calendar)
            => new Schedule(effectiveDate, maturityDate, new Period(frequency), calendar,
                            BusinessDayConvention.Unadjusted, BusinessDayConvention.Unadjusted,
                            DateGeneration.Rule.Backward, false);

        private static object[,] BuildCashflowTable(Leg leg, YieldTermStructureHandle curve)
        {
            int n      = (int)leg.Count;
            var result = new object[n, 4];
            for (int i = 0; i < n; i++)
            {
                var    cf     = leg[i];
                var    date   = cf.date();
                double amount = cf.amount();
                double df     = curve.currentLink().discount(date);
                result[i, 0] = QLHelper.ToExcelDate(date);
                result[i, 1] = amount;
                result[i, 2] = df;
                result[i, 3] = amount * df;
            }
            return result;
        }

        private static Schedule MakeFRNSchedule(Date issueQL, Date matQL,
                                                 Frequency freq, Calendar calendar)
            => new Schedule(issueQL, matQL, new Period(freq), calendar,
                            BusinessDayConvention.ModifiedFollowing,
                            BusinessDayConvention.ModifiedFollowing,
                            DateGeneration.Rule.Backward, false);

        // ─── Fixed-rate bond ──────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_FixedBondPrice",
                       Category = "QuantLib — Bonds",
                       Description = "Clean price, dirty price and accrued interest of a fixed-rate bond.\n" +
                                     "curve: a flat rate (e.g. 0.045) OR a handle from QL_BuildSwapCurve.\n" +
                                     "Returns 1×3 array: [CleanPrice, DirtyPrice, AccruedInterest].")]
        public static object QL_FixedBondPrice(
            [ExcelArgument(Description = "Face amount (e.g. 100)")] double faceAmount,
            [ExcelArgument(Description = "Annual coupon rate (decimal, e.g. 0.05)")] double couponRate,
            [ExcelArgument(Description = "Issue date (Excel date)")] double issueDate,
            [ExcelArgument(Description = "Maturity date (Excel date)")] double maturityDate,
            [ExcelArgument(Description = "Discount curve: flat rate (e.g. 0.045) or curve handle")] object curve,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Coupon frequency (default Semiannual)")] string frequency = "Semiannual",
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365",
            [ExcelArgument(Description = "Settlement days (default 3)")] int settlementDays = 3)
        {
            try
            {
                var evalQL    = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var discCurve = ResolveCurve(curve, evalQL);
                var freq      = QLHelper.ParseFrequency(frequency);
                var dc        = QLHelper.ParseDayCounter(dayCounter);
                var schedule  = MakeSchedule(QLHelper.ToQLDate(issueDate),
                                             QLHelper.ToQLDate(maturityDate), freq, new TARGET());
                var bond      = new FixedRateBond(settlementDays, faceAmount, schedule,
                                    QLHelper.ToDoubleVector(new[] { couponRate }), dc);
                bond.setPricingEngine(new DiscountingBondEngine(discCurve));
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
                                     "curve: a flat rate (e.g. 0.045) OR a handle from QL_BuildSwapCurve.\n" +
                                     "Returns columns: [Payment Date, Amount, Discount Factor, PV]. Array formula.")]
        public static object QL_FixedBondCashflows(
            [ExcelArgument(Description = "Face amount (e.g. 100)")] double faceAmount,
            [ExcelArgument(Description = "Annual coupon rate (decimal, e.g. 0.05)")] double couponRate,
            [ExcelArgument(Description = "Issue date (Excel date)")] double issueDate,
            [ExcelArgument(Description = "Maturity date (Excel date)")] double maturityDate,
            [ExcelArgument(Description = "Discount curve: flat rate or curve handle")] object curve,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Coupon frequency (default Semiannual)")] string frequency = "Semiannual",
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365",
            [ExcelArgument(Description = "Settlement days (default 3)")] int settlementDays = 3)
        {
            try
            {
                var evalQL    = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var discCurve = ResolveCurve(curve, evalQL);
                var freq      = QLHelper.ParseFrequency(frequency);
                var dc        = QLHelper.ParseDayCounter(dayCounter);
                var schedule  = MakeSchedule(QLHelper.ToQLDate(issueDate),
                                             QLHelper.ToQLDate(maturityDate), freq, new TARGET());
                var bond      = new FixedRateBond(settlementDays, faceAmount, schedule,
                                    QLHelper.ToDoubleVector(new[] { couponRate }), dc);
                bond.setPricingEngine(new DiscountingBondEngine(discCurve));
                return BuildCashflowTable(bond.cashflows(), discCurve);
            }
            catch (Exception ex) { return new object[,] { { $"#QL_ERR: {ex.Message}" } }; }
        }

        // ─── Floating-rate note ───────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_FRNPrice",
                       Category = "QuantLib — Bonds",
                       Description = "Clean price, dirty price and accrued interest of a floating-rate note (Euribor6M).\n" +
                                     "curve: a flat rate (e.g. 0.045) OR a handle from QL_BuildSwapCurve.\n" +
                                     "Returns 1×3 array: [CleanPrice, DirtyPrice, AccruedInterest].")]
        public static object QL_FRNPrice(
            [ExcelArgument(Description = "Face amount (e.g. 100)")] double faceAmount,
            [ExcelArgument(Description = "Spread over Euribor6M (decimal, e.g. 0.005 for +50bp)")] double spread,
            [ExcelArgument(Description = "Issue date (Excel date)")] double issueDate,
            [ExcelArgument(Description = "Maturity date (Excel date)")] double maturityDate,
            [ExcelArgument(Description = "Discount / forecast curve: flat rate or curve handle")] object curve,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Coupon frequency (default Semiannual)")] string frequency = "Semiannual",
            [ExcelArgument(Description = "Settlement days (default 3)")] int settlementDays = 3)
        {
            try
            {
                var evalQL    = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var discCurve = ResolveCurve(curve, evalQL);
                var freq      = QLHelper.ParseFrequency(frequency);
                var bond      = BuildFRN(faceAmount, spread, issueDate, maturityDate,
                                         evalQL, freq, settlementDays, discCurve);
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
                       Description = "Cashflow schedule of a floating-rate note (Euribor6M).\n" +
                                     "curve: a flat rate (e.g. 0.045) OR a handle from QL_BuildSwapCurve.\n" +
                                     "Returns columns: [Payment Date, Amount, Discount Factor, PV]. Array formula.")]
        public static object QL_FRNCashflows(
            [ExcelArgument(Description = "Face amount (e.g. 100)")] double faceAmount,
            [ExcelArgument(Description = "Spread over Euribor6M (decimal)")] double spread,
            [ExcelArgument(Description = "Issue date (Excel date)")] double issueDate,
            [ExcelArgument(Description = "Maturity date (Excel date)")] double maturityDate,
            [ExcelArgument(Description = "Discount / forecast curve: flat rate or curve handle")] object curve,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Coupon frequency (default Semiannual)")] string frequency = "Semiannual",
            [ExcelArgument(Description = "Settlement days (default 3)")] int settlementDays = 3)
        {
            try
            {
                var evalQL    = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var discCurve = ResolveCurve(curve, evalQL);
                var freq      = QLHelper.ParseFrequency(frequency);
                var bond      = BuildFRN(faceAmount, spread, issueDate, maturityDate,
                                         evalQL, freq, settlementDays, discCurve);
                return BuildCashflowTable(bond.cashflows(), discCurve);
            }
            catch (Exception ex) { return new object[,] { { $"#QL_ERR: {ex.Message}" } }; }
        }

        // ─── Internal FRN builder ─────────────────────────────────────────────────

        private static FloatingRateBond BuildFRN(
            double faceAmount, double spread,
            double issueDateSerial, double maturitySerial,
            Date evalQL, Frequency freq, int settlementDays,
            YieldTermStructureHandle discCurve)
        {
            var issueQL  = QLHelper.ToQLDate(issueDateSerial);
            var matQL    = QLHelper.ToQLDate(maturitySerial);
            var index    = new Euribor6M(discCurve);
            var calendar = new TARGET();
            var schedule = MakeFRNSchedule(issueQL, matQL, freq, calendar);

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
                issueQL);

            bond.setPricingEngine(new DiscountingBondEngine(discCurve));
            return bond;
        }
    }
}
