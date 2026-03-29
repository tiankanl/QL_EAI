/*
 * ExoticOptions.cs — Lookback, compound, chooser, exchange, and performance options.
 * Mirrors ql/instruments/lookbackoption.*, compoundoption.*, etc.
 */

using System;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.Options
{
    public static class ExoticOptions
    {
        // ─── Lookback ─────────────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_LookbackFixedOption",
                       Category = "QuantLib — Exotic Options",
                       Description = "Price a continuous fixed-strike lookback option (analytic formula).\n" +
                                     "Payoff: max(S_max - K, 0) for a call or max(K - S_min, 0) for a put.")]
        public static object QL_LookbackFixedOption(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Fixed strike")] double strike,
            [ExcelArgument(Description = "Running minimum/maximum of underlying so far (use spot if no history)")] double minMax,
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
                var option  = new ContinuousFixedLookbackOption(
                                  minMax, payoff,
                                  new EuropeanExercise(QLHelper.ToQLDate(maturityDate)));
                option.setPricingEngine(new AnalyticContinuousFixedLookbackEngine(process));
                return option.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_LookbackFloatingOption",
                       Category = "QuantLib — Exotic Options",
                       Description = "Price a continuous floating-strike lookback option (analytic formula).\n" +
                                     "Payoff: S_T - S_min (call) or S_max - S_T (put).")]
        public static object QL_LookbackFloatingOption(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Running minimum (for call) or maximum (for put) so far (use spot if no history)")] double minMax,
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
                var payoff  = new FloatingTypePayoff(QLHelper.ParseOptionType(optionType));
                var option  = new ContinuousFloatingLookbackOption(
                                  minMax, payoff,
                                  new EuropeanExercise(QLHelper.ToQLDate(maturityDate)));
                option.setPricingEngine(new AnalyticContinuousFloatingLookbackEngine(process));
                return option.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Compound option ──────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_CompoundOption",
                       Category = "QuantLib — Exotic Options",
                       Description = "Price a compound option (option on an option) using the analytic formula.\n" +
                                     "Outer type: type of the compound option itself.\n" +
                                     "Inner type: type of the underlying option.")]
        public static object QL_CompoundOption(
            [ExcelArgument(Description = "Outer option type: \"call\" or \"put\"")] string outerType,
            [ExcelArgument(Description = "Inner option type: \"call\" or \"put\"")] string innerType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Inner strike (strike on the underlying option)")] double innerStrike,
            [ExcelArgument(Description = "Outer strike (premium paid for underlying option)")] double outerStrike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Outer expiry date (compound option expiry)")] double outerExpiry,
            [ExcelArgument(Description = "Inner expiry date (underlying option expiry)")] double innerExpiry)
        {
            try
            {
                var evalQL  = QLHelper.ToQLDate(evalDate);
                var process = QLHelper.BuildBSMProcess(spot, dividendYield, riskFreeRate, volatility, evalQL);
                var outerPayoff = new PlainVanillaPayoff(QLHelper.ParseOptionType(outerType), outerStrike);
                var innerPayoff = new PlainVanillaPayoff(QLHelper.ParseOptionType(innerType), innerStrike);
                var option = new CompoundOption(
                    outerPayoff, new EuropeanExercise(QLHelper.ToQLDate(outerExpiry)),
                    innerPayoff, new EuropeanExercise(QLHelper.ToQLDate(innerExpiry)));
                option.setPricingEngine(new AnalyticCompoundOptionEngine(process));
                return option.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Chooser option ───────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_SimpleChooserOption",
                       Category = "QuantLib — Exotic Options",
                       Description = "Price a simple chooser option (analytic formula).\n" +
                                     "At chooserDate the holder picks the better of a call or a put with the same strike and expiry.")]
        public static object QL_SimpleChooserOption(
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike (same for both underlying call and put)")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Choice date (when holder picks call or put)")] double chooserDate,
            [ExcelArgument(Description = "Underlying option expiry (must be >= chooserDate)")] double maturityDate)
        {
            try
            {
                var evalQL  = QLHelper.ToQLDate(evalDate);
                var process = QLHelper.BuildBSMProcess(spot, dividendYield, riskFreeRate, volatility, evalQL);
                var option  = new SimpleChooserOption(
                                  QLHelper.ToQLDate(chooserDate), strike,
                                  new EuropeanExercise(QLHelper.ToQLDate(maturityDate)));
                option.setPricingEngine(new AnalyticSimpleChooserEngine(process));
                return option.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Exchange option (Margrabe) ───────────────────────────────────────────

        [ExcelFunction(Name = "QL_ExchangeOption",
                       Category = "QuantLib — Exotic Options",
                       Description = "Price a European exchange option (Margrabe formula).\n" +
                                     "Payoff: max(S1 - S2, 0), where S1 and S2 are two correlated assets.")]
        public static object QL_ExchangeOption(
            [ExcelArgument(Description = "Spot price of asset 1 (asset to acquire)")] double spot1,
            [ExcelArgument(Description = "Spot price of asset 2 (asset to give up)")] double spot2,
            [ExcelArgument(Description = "Dividend yield of asset 1")] double divYield1,
            [ExcelArgument(Description = "Dividend yield of asset 2")] double divYield2,
            [ExcelArgument(Description = "Volatility of asset 1")] double vol1,
            [ExcelArgument(Description = "Volatility of asset 2")] double vol2,
            [ExcelArgument(Description = "Correlation between the two assets")] double correlation,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate)
        {
            try
            {
                var evalQL     = QLHelper.ToQLDate(evalDate);
                Settings.instance().setEvaluationDate(evalQL);
                var matQL      = QLHelper.ToQLDate(maturityDate);
                var dayCounter = new Actual365Fixed();
                var calendar   = new TARGET();

                var p1 = new BlackScholesMertonProcess(
                             new QuoteHandle(new SimpleQuote(spot1)),
                             new YieldTermStructureHandle(new FlatForward(evalQL, divYield1, dayCounter)),
                             new YieldTermStructureHandle(new FlatForward(evalQL, riskFreeRate, dayCounter)),
                             new BlackVolTermStructureHandle(new BlackConstantVol(evalQL, calendar, vol1, dayCounter)));
                var p2 = new BlackScholesMertonProcess(
                             new QuoteHandle(new SimpleQuote(spot2)),
                             new YieldTermStructureHandle(new FlatForward(evalQL, divYield2, dayCounter)),
                             new YieldTermStructureHandle(new FlatForward(evalQL, riskFreeRate, dayCounter)),
                             new BlackVolTermStructureHandle(new BlackConstantVol(evalQL, calendar, vol2, dayCounter)));

                var option = new MargrabeOption(1, 1, new EuropeanExercise(matQL));
                option.setPricingEngine(new AnalyticEuropeanMargrabeEngine(p1, p2, correlation));
                return option.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Cliquet option ───────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_CliquetOption",
                       Category = "QuantLib — Exotic Options",
                       Description = "Price a cliquet (ratchet) option using the analytic forward-start formula.\n" +
                                     "Provide reset dates as an Excel range. The last reset date is the expiry.")]
        public static object QL_CliquetOption(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike as a fraction of spot at each reset (e.g. 1.0 for ATM)")] double moneyness,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Reset dates (Excel date serials); last entry = expiry")] object resetDates)
        {
            try
            {
                var evalQL  = QLHelper.ToQLDate(evalDate);
                var process = QLHelper.BuildBSMProcess(spot, dividendYield, riskFreeRate, volatility, evalQL);
                var dates   = QLHelper.ToDateVector(resetDates);
                var payoff  = new PercentageStrikePayoff(QLHelper.ParseOptionType(optionType), moneyness);
                var option  = new CliquetOption(payoff, new EuropeanExercise(dates[(int)dates.Count - 1]), dates);
                option.setPricingEngine(new AnalyticCliquetEngine(process));
                return option.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }
    }
}
