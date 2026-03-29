/*
 * BarrierOptions.cs — Single-barrier and double-barrier option pricing.
 * Mirrors ql/instruments/barrieroption.* and ql/instruments/doublebarrieroption.*
 */

using System;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.Options
{
    public static class BarrierOptions
    {
        [ExcelFunction(Name = "QL_BarrierOption",
                       Category = "QuantLib — Barrier Options",
                       Description = "Price a single-barrier option using the analytic BSM formula.\n" +
                                     "BarrierType: DownIn | DownOut | UpIn | UpOut.")]
        public static object QL_BarrierOption(
            [ExcelArgument(Description = "Barrier type: \"DownIn\", \"DownOut\", \"UpIn\", or \"UpOut\"")] string barrierType,
            [ExcelArgument(Description = "Barrier level")] double barrier,
            [ExcelArgument(Description = "Rebate paid if barrier is triggered or not reached (0 = none)")] double rebate,
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
                var bt      = QLHelper.ParseBarrierType(barrierType);
                var option  = new BarrierOption(bt, barrier, rebate, payoff,
                                                new EuropeanExercise(QLHelper.ToQLDate(maturityDate)));
                option.setPricingEngine(new AnalyticBarrierEngine(process));
                return option.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BarrierOptionDelta",
                       Category = "QuantLib — Barrier Options",
                       Description = "Delta of a single-barrier option.")]
        public static object QL_BarrierOptionDelta(
            [ExcelArgument(Description = "Barrier type: \"DownIn\", \"DownOut\", \"UpIn\", or \"UpOut\"")] string barrierType,
            [ExcelArgument(Description = "Barrier level")] double barrier,
            [ExcelArgument(Description = "Rebate")] double rebate,
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
                var option  = new BarrierOption(QLHelper.ParseBarrierType(barrierType), barrier, rebate, payoff,
                                                new EuropeanExercise(QLHelper.ToQLDate(maturityDate)));
                option.setPricingEngine(new AnalyticBarrierEngine(process));
                return option.delta();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_DoubleBarrierOption",
                       Category = "QuantLib — Barrier Options",
                       Description = "Price a double-barrier option (analytic BSM formula).\n" +
                                     "BarrierType: KnockIn | KnockOut | KIKO | KOKI.\n" +
                                     "lowerBarrier < strike < upperBarrier (for standard structures).")]
        public static object QL_DoubleBarrierOption(
            [ExcelArgument(Description = "Barrier type: \"KnockIn\", \"KnockOut\", \"KIKO\", or \"KOKI\"")] string barrierType,
            [ExcelArgument(Description = "Lower barrier level")] double lowerBarrier,
            [ExcelArgument(Description = "Upper barrier level")] double upperBarrier,
            [ExcelArgument(Description = "Rebate (0 = none)")] double rebate,
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
                var bt      = QLHelper.ParseDoubleBarrierType(barrierType);
                var option  = new DoubleBarrierOption(bt, lowerBarrier, upperBarrier, rebate, payoff,
                                                      new EuropeanExercise(QLHelper.ToQLDate(maturityDate)));
                option.setPricingEngine(new AnalyticDoubleBarrierEngine(process));
                return option.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BinaryBarrierOption",
                       Category = "QuantLib — Barrier Options",
                       Description = "Price a binary (cash-or-nothing) barrier option using analytic formula.\n" +
                                     "BarrierType: DownIn | DownOut | UpIn | UpOut.")]
        public static object QL_BinaryBarrierOption(
            [ExcelArgument(Description = "Barrier type: \"DownIn\", \"DownOut\", \"UpIn\", or \"UpOut\"")] string barrierType,
            [ExcelArgument(Description = "Barrier level")] double barrier,
            [ExcelArgument(Description = "Cash rebate / payoff amount")] double cashPayoff,
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
                var payoff  = new CashOrNothingPayoff(QLHelper.ParseOptionType(optionType), strike, cashPayoff);
                var bt      = QLHelper.ParseBarrierType(barrierType);
                var option  = new BarrierOption(bt, barrier, 0.0, payoff,
                                                new EuropeanExercise(QLHelper.ToQLDate(maturityDate)));
                option.setPricingEngine(new AnalyticBinaryBarrierEngine(process));
                return option.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }
    }
}
