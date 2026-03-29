/*
 * AsianOptions.cs — Asian (average-price / average-strike) option pricing.
 * Mirrors ql/instruments/asianoption.* and ql/pricingengines/asian/*
 */

using System;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.Options
{
    public static class AsianOptions
    {
        [ExcelFunction(Name = "QL_ContinuousGeometricAsian",
                       Category = "QuantLib — Asian Options",
                       Description = "Price a continuous geometric average-price Asian option (analytic formula).")]
        public static object QL_ContinuousGeometricAsian(
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
                var evalQL  = QLHelper.ToQLDate(evalDate);
                var process = QLHelper.BuildBSMProcess(spot, dividendYield, riskFreeRate, volatility, evalQL);
                var payoff  = new PlainVanillaPayoff(QLHelper.ParseOptionType(optionType), strike);
                var option  = new ContinuousAveragingAsianOption(
                                  Average.Type.Geometric, payoff,
                                  new EuropeanExercise(QLHelper.ToQLDate(maturityDate)));
                option.setPricingEngine(
                    new AnalyticContinuousGeometricAveragePriceAsianEngine(process));
                return option.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_DiscreteGeometricAsian",
                       Category = "QuantLib — Asian Options",
                       Description = "Price a discrete geometric average-price Asian option (analytic formula).\n" +
                                     "Provide fixing dates as an Excel range (column or row of date serial numbers).")]
        public static object QL_DiscreteGeometricAsian(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Range of fixing dates (Excel date serials)")] object fixingDates,
            [ExcelArgument(Description = "Running sum of past fixings (0 if no past fixings)")] double runningSum = 0.0,
            [ExcelArgument(Description = "Number of past fixings already accumulated (0 if none)")] int pastFixings = 0)
        {
            try
            {
                var evalQL  = QLHelper.ToQLDate(evalDate);
                var process = QLHelper.BuildBSMProcess(spot, dividendYield, riskFreeRate, volatility, evalQL);
                var payoff  = new PlainVanillaPayoff(QLHelper.ParseOptionType(optionType), strike);
                var dates   = QLHelper.ToDateVector(fixingDates);
                var option  = new DiscreteAveragingAsianOption(
                                  Average.Type.Geometric, runningSum, (uint)pastFixings,
                                  dates, payoff,
                                  new EuropeanExercise(QLHelper.ToQLDate(maturityDate)));
                option.setPricingEngine(
                    new AnalyticDiscreteGeometricAveragePriceAsianEngine(process));
                return option.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_DiscreteArithmeticAsianMC",
                       Category = "QuantLib — Asian Options",
                       Description = "Price a discrete arithmetic average-price Asian option via Monte Carlo.\n" +
                                     "Uses geometric average as a control variate for variance reduction.")]
        public static object QL_DiscreteArithmeticAsianMC(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Range of fixing dates (Excel date serials)")] object fixingDates,
            [ExcelArgument(Description = "Number of MC samples (default 10000)")] int samples = 10000,
            [ExcelArgument(Description = "Random seed (default 42)")] int seed = 42)
        {
            try
            {
                var evalQL  = QLHelper.ToQLDate(evalDate);
                var process = QLHelper.BuildBSMProcess(spot, dividendYield, riskFreeRate, volatility, evalQL);
                var payoff  = new PlainVanillaPayoff(QLHelper.ParseOptionType(optionType), strike);
                var dates   = QLHelper.ToDateVector(fixingDates);
                var option  = new DiscreteAveragingAsianOption(
                                  Average.Type.Arithmetic, 0.0, 0,
                                  dates, payoff,
                                  new EuropeanExercise(QLHelper.ToQLDate(maturityDate)));
                option.setPricingEngine(
                    new MCPRDiscreteArithmeticAPEngine(process,
                        brownianBridge: true, antitheticVariate: true,
                        controlVariate: true, requiredSamples: samples,
                        requiredTolerance: 0, maxSamples: samples * 2, seed: seed));
                return option.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }
    }
}
