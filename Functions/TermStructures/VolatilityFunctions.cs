/*
 * VolatilityFunctions.cs — Black volatility surfaces, SABR, and related utilities.
 * Mirrors ql/termstructures/volatility/equityfx/* and ql/termstructures/volatility/sabr*.
 */

using System;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.TermStructures
{
    public static class VolatilityFunctions
    {
        [ExcelFunction(Name = "QL_BlackConstantVol",
                       Category = "QuantLib — Volatility",
                       Description = "Query a flat (constant) Black volatility surface at a given date and strike.")]
        public static object QL_BlackConstantVol(
            [ExcelArgument(Description = "Flat volatility (decimal, e.g. 0.20)")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Expiry date to query")] double expiryDate,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365")
        {
            try
            {
                var evalQL   = QLHelper.ToQLDate(evalDate);
                var expiryQL = QLHelper.ToQLDate(expiryDate);
                Settings.instance().setEvaluationDate(evalQL);
                var dc       = QLHelper.ParseDayCounter(dayCounter);
                var surface  = new BlackConstantVol(evalQL, new TARGET(), volatility, dc);
                return surface.blackVol(expiryQL, strike);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BlackVariance",
                       Category = "QuantLib — Volatility",
                       Description = "Integrated Black variance (σ² × T) for a flat volatility surface.")]
        public static object QL_BlackVariance(
            [ExcelArgument(Description = "Flat volatility (decimal)")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Expiry date")] double expiryDate,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Day counter (default Actual365)")] string dayCounter = "Actual365")
        {
            try
            {
                var evalQL   = QLHelper.ToQLDate(evalDate);
                var expiryQL = QLHelper.ToQLDate(expiryDate);
                Settings.instance().setEvaluationDate(evalQL);
                var dc       = QLHelper.ParseDayCounter(dayCounter);
                var surface  = new BlackConstantVol(evalQL, new TARGET(), volatility, dc);
                return surface.blackVariance(expiryQL, strike);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_SABRVol",
                       Category = "QuantLib — Volatility",
                       Description = "SABR implied Black volatility (Hagan et al. 2002).\n" +
                                     "Parameters: forward F, strike K, expiry T in years,\n" +
                                     "alpha (vol of vol level), beta (CEV exponent 0-1),\n" +
                                     "nu (vol of vol), rho (spot-vol correlation).")]
        public static object QL_SABRVol(
            [ExcelArgument(Description = "Forward price / rate F")] double forward,
            [ExcelArgument(Description = "Strike K")] double strike,
            [ExcelArgument(Description = "Time to expiry T (in years)")] double timeToExpiry,
            [ExcelArgument(Description = "SABR alpha (initial vol level, > 0)")] double alpha,
            [ExcelArgument(Description = "SABR beta (CEV exponent, 0 ≤ β ≤ 1)")] double beta,
            [ExcelArgument(Description = "SABR nu (vol of vol, ≥ 0)")] double nu,
            [ExcelArgument(Description = "SABR rho (spot-vol correlation, −1 < ρ < 1)")] double rho)
        {
            try
            {
                return NQuantLibc.sabrVolatility(strike, forward, timeToExpiry, alpha, beta, nu, rho);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_ShiftedSABRVol",
                       Category = "QuantLib — Volatility",
                       Description = "Shifted-SABR implied Black vol — useful for negative-rate environments.\n" +
                                     "The shift parameter moves the backbone so that F+shift and K+shift are used.")]
        public static object QL_ShiftedSABRVol(
            [ExcelArgument(Description = "Forward price / rate F")] double forward,
            [ExcelArgument(Description = "Strike K")] double strike,
            [ExcelArgument(Description = "Time to expiry T (in years)")] double timeToExpiry,
            [ExcelArgument(Description = "SABR alpha")] double alpha,
            [ExcelArgument(Description = "SABR beta")] double beta,
            [ExcelArgument(Description = "SABR nu")] double nu,
            [ExcelArgument(Description = "SABR rho")] double rho,
            [ExcelArgument(Description = "Shift (additive, e.g. 0.03 for 3%)")] double shift)
        {
            try
            {
                return NQuantLibc.shiftedSabrVolatility(strike, forward, timeToExpiry, alpha, beta, nu, rho, shift);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BlackFormula",
                       Category = "QuantLib — Volatility",
                       Description = "Black-76 option price (used for caplets, swaptions, futures options).\n" +
                                     "Inputs are forward F, strike K, stdDev = σ√T, discount P(0,T).")]
        public static object QL_BlackFormula(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Strike K")] double strike,
            [ExcelArgument(Description = "Forward price / rate F")] double forward,
            [ExcelArgument(Description = "Standard deviation σ√T")] double stdDev,
            [ExcelArgument(Description = "Discount factor P(0,T) (default 1.0)")] double discount = 1.0)
        {
            try
            {
                var otype = QLHelper.ParseOptionType(optionType);
                return NQuantLibc.blackFormula(otype, strike, forward, stdDev, discount);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BlackFormulaDelta",
                       Category = "QuantLib — Volatility",
                       Description = "Delta from the Black-76 formula: ∂V/∂F.")]
        public static object QL_BlackFormulaDelta(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Strike K")] double strike,
            [ExcelArgument(Description = "Forward F")] double forward,
            [ExcelArgument(Description = "Standard deviation σ√T")] double stdDev,
            [ExcelArgument(Description = "Discount factor P(0,T) (default 1.0)")] double discount = 1.0)
        {
            try
            {
                // Black-76 delta = ∂V/∂F = discount * N(d1) for call, discount * (N(d1)-1) for put
                var otype = QLHelper.ParseOptionType(optionType);
                double d1 = stdDev > 0 ? (global::System.Math.Log(forward / strike) / stdDev + stdDev / 2.0) : 0;
                double nd1 = new CumulativeNormalDistribution().call(d1);
                return discount * (otype == Option.Type.Call ? nd1 : nd1 - 1.0);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BlackFormulaImpliedStdDev",
                       Category = "QuantLib — Volatility",
                       Description = "Implied standard deviation σ√T from a Black-76 price.")]
        public static object QL_BlackFormulaImpliedStdDev(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Strike K")] double strike,
            [ExcelArgument(Description = "Forward F")] double forward,
            [ExcelArgument(Description = "Market option price")] double optionPrice,
            [ExcelArgument(Description = "Discount factor P(0,T) (default 1.0)")] double discount = 1.0,
            [ExcelArgument(Description = "Displacement / shift (default 0)")] double displacement = 0.0,
            [ExcelArgument(Description = "Accuracy (default 1e-6)")] double accuracy = 1e-6,
            [ExcelArgument(Description = "Max iterations (default 300)")] int maxIter = 300,
            [ExcelArgument(Description = "Initial vol guess (default 0.5)")] double guess = 0.5)
        {
            try
            {
                var otype = QLHelper.ParseOptionType(optionType);
                return NQuantLibc.blackFormulaImpliedStdDev(
                    otype, strike, forward, optionPrice, discount, displacement, guess, accuracy, (uint)maxIter);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }
    }
}
