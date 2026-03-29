/*
 * YieldCurveFunctions.cs — Yield curve queries: discount factors, zero rates, forward rates, par rates.
 * Mirrors ql/termstructures/yield/*.
 */

using System;
using SM = System.Math;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.TermStructures
{
    public static class YieldCurveFunctions
    {
        [ExcelFunction(Name = "QL_DiscountFactor",
                       Category = "QuantLib — Yield Curves",
                       Description = "Discount factor P(0,T) from a flat continuously-compounded risk-free rate.")]
        public static object QL_DiscountFactor(
            [ExcelArgument(Description = "Risk-free rate (decimal)")] double riskFreeRate,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Target date (Excel date)")] double targetDate,
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365")
        {
            try
            {
                var evalQL  = QLHelper.ToQLDate(evalDate);
                var targQL  = QLHelper.ToQLDate(targetDate);
                Settings.instance().setEvaluationDate(evalQL);
                var dc      = QLHelper.ParseDayCounter(dayCounter);
                var curve   = new FlatForward(evalQL, riskFreeRate, dc);
                return curve.discount(targQL);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_ZeroRate",
                       Category = "QuantLib — Yield Curves",
                       Description = "Zero (spot) rate implied by a flat forward curve.")]
        public static object QL_ZeroRate(
            [ExcelArgument(Description = "Flat forward rate (decimal)")] double flatRate,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Target date")] double targetDate,
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365",
            [ExcelArgument(Description = "Compounding (default Continuous)")] string compounding = "Continuous",
            [ExcelArgument(Description = "Frequency (default Annual)")] string frequency = "Annual")
        {
            try
            {
                var evalQL = QLHelper.ToQLDate(evalDate);
                var targQL = QLHelper.ToQLDate(targetDate);
                Settings.instance().setEvaluationDate(evalQL);
                var dc     = QLHelper.ParseDayCounter(dayCounter);
                var comp   = QLHelper.ParseCompounding(compounding);
                var freq   = QLHelper.ParseFrequency(frequency);
                var curve  = new FlatForward(evalQL, flatRate, dc);
                return curve.zeroRate(targQL, dc, comp, freq).rate();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_ForwardRate",
                       Category = "QuantLib — Yield Curves",
                       Description = "Instantaneous or period forward rate between two dates from a flat yield curve.")]
        public static object QL_ForwardRate(
            [ExcelArgument(Description = "Flat forward rate (decimal)")] double flatRate,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Forward period start date")] double date1,
            [ExcelArgument(Description = "Forward period end date")] double date2,
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365",
            [ExcelArgument(Description = "Compounding (default Continuous)")] string compounding = "Continuous",
            [ExcelArgument(Description = "Frequency (default Annual)")] string frequency = "Annual")
        {
            try
            {
                var evalQL = QLHelper.ToQLDate(evalDate);
                var d1     = QLHelper.ToQLDate(date1);
                var d2     = QLHelper.ToQLDate(date2);
                Settings.instance().setEvaluationDate(evalQL);
                var dc     = QLHelper.ParseDayCounter(dayCounter);
                var comp   = QLHelper.ParseCompounding(compounding);
                var freq   = QLHelper.ParseFrequency(frequency);
                var curve  = new FlatForward(evalQL, flatRate, dc);
                return curve.forwardRate(d1, d2, dc, comp, freq).rate();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_ParRate",
                       Category = "QuantLib — Yield Curves",
                       Description = "Par coupon rate for a fixed-rate bond priced at par (100) from a flat yield curve.")]
        public static object QL_ParRate(
            [ExcelArgument(Description = "Flat discount rate (decimal)")] double flatRate,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Coupon frequency (default Semiannual)")] string frequency = "Semiannual",
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365")
        {
            try
            {
                var evalQL = QLHelper.ToQLDate(evalDate);
                var matQL  = QLHelper.ToQLDate(maturityDate);
                Settings.instance().setEvaluationDate(evalQL);
                var dc     = QLHelper.ParseDayCounter(dayCounter);
                var freq   = QLHelper.ParseFrequency(frequency);
                var curve  = new FlatForward(evalQL, flatRate, dc);
                var handle = new YieldTermStructureHandle(curve);

                // par rate = (1 - P(T)) / sum(P(Ti) * delta_i)
                // Build coupon dates and compute the annuity analytically.
                double period = 1.0 / (int)freq;
                double t      = 0.0;
                double tMax   = dc.yearFraction(evalQL, matQL);
                double annuity = 0.0;
                double lastPeriod = period;
                while (t + period <= tMax + 1e-8)
                {
                    t += period;
                    var ti = evalQL + (int)SM.Round(t * 365);
                    annuity += curve.discount(ti) * period;
                }
                double finalDiscount = curve.discount(matQL);
                return (1.0 - finalDiscount) / annuity;
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_ImpliedForwardRate",
                       Category = "QuantLib — Yield Curves",
                       Description = "Forward rate implied from two spot zero rates (bootstrapped from flat curves).")]
        public static object QL_ImpliedForwardRate(
            [ExcelArgument(Description = "Near-date zero rate (decimal)")] double nearRate,
            [ExcelArgument(Description = "Far-date zero rate (decimal)")] double farRate,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Near date")] double nearDate,
            [ExcelArgument(Description = "Far date")] double farDate,
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365")
        {
            try
            {
                var evalQL  = QLHelper.ToQLDate(evalDate);
                var nearQL  = QLHelper.ToQLDate(nearDate);
                var farQL   = QLHelper.ToQLDate(farDate);
                Settings.instance().setEvaluationDate(evalQL);
                var dc      = QLHelper.ParseDayCounter(dayCounter);
                double t1   = dc.yearFraction(evalQL, nearQL);
                double t2   = dc.yearFraction(evalQL, farQL);
                // Continuous compounding: P1 = exp(-r1*t1), P2 = exp(-r2*t2)
                double p1   = SM.Exp(-nearRate * t1);
                double p2   = SM.Exp(-farRate  * t2);
                return -SM.Log(p2 / p1) / (t2 - t1);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_RateConvert",
                       Category = "QuantLib — Yield Curves",
                       Description = "Convert a rate between compounding conventions / frequencies.\n" +
                                     "e.g. convert a semiannually-compounded rate to continuously-compounded.")]
        public static object QL_RateConvert(
            [ExcelArgument(Description = "Input rate (decimal)")] double rate,
            [ExcelArgument(Description = "Input compounding (e.g. Compounded)")] string fromCompounding,
            [ExcelArgument(Description = "Input frequency (e.g. Semiannual)")] string fromFrequency,
            [ExcelArgument(Description = "Output compounding (e.g. Continuous)")] string toCompounding,
            [ExcelArgument(Description = "Output frequency (e.g. Annual)")] string toFrequency,
            [ExcelArgument(Description = "Time period in years")] double years,
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365")
        {
            try
            {
                var dc      = QLHelper.ParseDayCounter(dayCounter);
                var fromC   = QLHelper.ParseCompounding(fromCompounding);
                var fromF   = QLHelper.ParseFrequency(fromFrequency);
                var toC     = QLHelper.ParseCompounding(toCompounding);
                var toF     = QLHelper.ParseFrequency(toFrequency);
                var ir      = new InterestRate(rate, dc, fromC, fromF);
                return ir.equivalentRate(toC, toF, years).rate();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }
    }
}
