/*
 * AmericanOptions.cs — American vanilla options.
 * Engines: Binomial CRR, Barone-Adesi-Whaley, Bjerksund-Stensland.
 * Mirrors ql/pricingengines/vanilla/*american*.
 */

using System;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.Options
{
    public static class AmericanOptions
    {
        private static double PriceAmerican(
            string optionType, double spot, double strike,
            double riskFreeRate, double dividendYield, double volatility,
            double evalSerial, double maturitySerial,
            Func<BlackScholesMertonProcess, PricingEngine> engineFactory)
        {
            var evalDate = QLHelper.ToQLDate(evalSerial);
            var matDate  = QLHelper.ToQLDate(maturitySerial);
            var process  = QLHelper.BuildBSMProcess(spot, dividendYield, riskFreeRate, volatility, evalDate);
            var payoff   = new PlainVanillaPayoff(QLHelper.ParseOptionType(optionType), strike);
            var exercise = new AmericanExercise(evalDate, matDate);
            var option   = new VanillaOption(payoff, exercise);
            option.setPricingEngine(engineFactory(process));
            return option.NPV();
        }

        [ExcelFunction(Name = "QL_AmericanBinomialCRR",
                       Category = "QuantLib — American Options",
                       Description = "Price an American vanilla option using the Binomial Cox-Ross-Rubinstein lattice.")]
        public static object QL_AmericanBinomialCRR(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Number of time steps (default 200)")] int timeSteps = 200)
        {
            try
            {
                return PriceAmerican(optionType, spot, strike, riskFreeRate, dividendYield,
                    volatility, evalDate, maturityDate,
                    p => new BinomialCRRVanillaEngine(p, (uint)timeSteps));
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_AmericanBinomialJR",
                       Category = "QuantLib — American Options",
                       Description = "Price an American vanilla option using the Binomial Jarrow-Rudd lattice.")]
        public static object QL_AmericanBinomialJR(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Number of time steps (default 200)")] int timeSteps = 200)
        {
            try
            {
                return PriceAmerican(optionType, spot, strike, riskFreeRate, dividendYield,
                    volatility, evalDate, maturityDate,
                    p => new BinomialJRVanillaEngine(p, (uint)timeSteps));
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_AmericanBAW",
                       Category = "QuantLib — American Options",
                       Description = "Price an American vanilla option using the Barone-Adesi-Whaley approximation (fast analytic).")]
        public static object QL_AmericanBAW(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate)
        {
            try
            {
                return PriceAmerican(optionType, spot, strike, riskFreeRate, dividendYield,
                    volatility, evalDate, maturityDate,
                    p => new BaroneAdesiWhaleyApproximationEngine(p));
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_AmericanBjerksund",
                       Category = "QuantLib — American Options",
                       Description = "Price an American vanilla option using the Bjerksund-Stensland approximation.")]
        public static object QL_AmericanBjerksund(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate)
        {
            try
            {
                return PriceAmerican(optionType, spot, strike, riskFreeRate, dividendYield,
                    volatility, evalDate, maturityDate,
                    p => new BjerksundStenslandApproximationEngine(p));
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_AmericanImpliedVol",
                       Category = "QuantLib — American Options",
                       Description = "Implied volatility for an American option using Binomial CRR (slower than BSM inversion).")]
        public static object QL_AmericanImpliedVol(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Market price")] double marketPrice,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Time steps for CRR (default 100)")] int timeSteps = 100,
            [ExcelArgument(Description = "Accuracy (default 1e-4)")] double accuracy = 1e-4,
            [ExcelArgument(Description = "Max iterations (default 200)")] int maxEval = 200)
        {
            try
            {
                var evalQL  = QLHelper.ToQLDate(evalDate);
                var matQL   = QLHelper.ToQLDate(maturityDate);
                var process = QLHelper.BuildBSMProcess(spot, dividendYield, riskFreeRate, 0.20, evalQL);
                var payoff  = new PlainVanillaPayoff(QLHelper.ParseOptionType(optionType), strike);
                var option  = new VanillaOption(payoff, new AmericanExercise(evalQL, matQL));
                return option.impliedVolatility(marketPrice, process, accuracy, (uint)maxEval, 1e-4, 4.0);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }
    }
}
