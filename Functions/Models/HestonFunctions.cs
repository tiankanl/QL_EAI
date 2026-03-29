/*
 * HestonFunctions.cs — Heston stochastic-volatility model option pricing.
 * Mirrors ql/models/equity/hestonmodel.* and ql/pricingengines/vanilla/analytichestonengine.*
 */

using System;
using SM = System.Math;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.Models
{
    public static class HestonFunctions
    {
        private static VanillaOption BuildHestonEuropean(
            string optionType, double spot, double strike,
            double riskFreeRate, double dividendYield,
            double v0, double kappa, double theta, double sigma, double rho,
            double evalSerial, double maturitySerial)
        {
            var evalQL = QLHelper.ToQLDate(evalSerial);
            var matQL  = QLHelper.ToQLDate(maturitySerial);
            Settings.instance().setEvaluationDate(evalQL);

            var dayCounter = new Actual365Fixed();
            var spotHandle = new QuoteHandle(new SimpleQuote(spot));
            var rtsHandle  = new YieldTermStructureHandle(
                                 new FlatForward(evalQL, riskFreeRate, dayCounter));
            var qtsHandle  = new YieldTermStructureHandle(
                                 new FlatForward(evalQL, dividendYield, dayCounter));

            var process = new HestonProcess(rtsHandle, qtsHandle, spotHandle,
                                            v0, kappa, theta, sigma, rho);
            var model   = new HestonModel(process);
            var engine  = new AnalyticHestonEngine(model);

            var payoff   = new PlainVanillaPayoff(QLHelper.ParseOptionType(optionType), strike);
            var exercise = new EuropeanExercise(matQL);
            var option   = new VanillaOption(payoff, exercise);
            option.setPricingEngine(engine);
            return option;
        }

        [ExcelFunction(Name = "QL_HestonPrice",
                       Category = "QuantLib — Heston Model",
                       Description = "European option price under the Heston stochastic-vol model. " +
                                     "v0=initial variance, kappa=mean-reversion speed, theta=long-run variance, " +
                                     "sigma=vol-of-vol, rho=spot/vol correlation.")]
        public static object QL_HestonPrice(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Initial variance v₀ (e.g. 0.04)")] double v0,
            [ExcelArgument(Description = "Mean-reversion speed κ")] double kappa,
            [ExcelArgument(Description = "Long-run variance θ")] double theta,
            [ExcelArgument(Description = "Vol-of-vol σ")] double sigma,
            [ExcelArgument(Description = "Spot-vol correlation ρ")] double rho,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate)
        {
            try
            {
                var option = BuildHestonEuropean(optionType, spot, strike, riskFreeRate, dividendYield,
                                                 v0, kappa, theta, sigma, rho, evalDate, maturityDate);
                return option.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_HestonDelta",
                       Category = "QuantLib — Heston Model",
                       Description = "Delta of a European option under the Heston model.")]
        public static object QL_HestonDelta(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "v₀")] double v0,
            [ExcelArgument(Description = "κ")] double kappa,
            [ExcelArgument(Description = "θ")] double theta,
            [ExcelArgument(Description = "σ")] double sigma,
            [ExcelArgument(Description = "ρ")] double rho,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate)
        {
            try
            {
                var option = BuildHestonEuropean(optionType, spot, strike, riskFreeRate, dividendYield,
                                                 v0, kappa, theta, sigma, rho, evalDate, maturityDate);
                return option.delta();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_HestonImpliedVol",
                       Category = "QuantLib — Heston Model",
                       Description = "Implied Black-Scholes volatility from a Heston model price (BSM inversion).")]
        public static object QL_HestonImpliedVol(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "v₀")] double v0,
            [ExcelArgument(Description = "κ")] double kappa,
            [ExcelArgument(Description = "θ")] double theta,
            [ExcelArgument(Description = "σ")] double sigma,
            [ExcelArgument(Description = "ρ")] double rho,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Accuracy (default 1e-6)")] double accuracy = 1e-6,
            [ExcelArgument(Description = "Max evaluations (default 300)")] int maxEval = 300)
        {
            try
            {
                var option      = BuildHestonEuropean(optionType, spot, strike, riskFreeRate, dividendYield,
                                                      v0, kappa, theta, sigma, rho, evalDate, maturityDate);
                var hestonPrice = option.NPV();

                // Now invert via BSM process (vol=0.2 placeholder; impliedVolatility searches for the right vol).
                var evalQL  = QLHelper.ToQLDate(evalDate);
                var process = QLHelper.BuildBSMProcess(spot, dividendYield, riskFreeRate, 0.20, evalQL);
                var payoff  = new PlainVanillaPayoff(QLHelper.ParseOptionType(optionType), strike);
                var eu      = new VanillaOption(payoff, new EuropeanExercise(QLHelper.ToQLDate(maturityDate)));
                return eu.impliedVolatility(hestonPrice, process, accuracy, (uint)maxEval, 1e-5, 5.0);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_HestonVolSmile",
                       Category = "QuantLib — Heston Model",
                       Description = "Implied volatility smile from Heston model at a fixed expiry for a range of moneyness.\n" +
                                     "Returns a vertical array of implied vols for log-moneyness from -nStdDev to +nStdDev.")]
        public static object QL_HestonVolSmile(
            [ExcelArgument(Description = "Spot / forward")] double spot,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "v₀")] double v0,
            [ExcelArgument(Description = "κ")] double kappa,
            [ExcelArgument(Description = "θ")] double theta,
            [ExcelArgument(Description = "σ")] double sigma,
            [ExcelArgument(Description = "ρ")] double rho,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate,
            [ExcelArgument(Description = "Number of strikes on each side of ATM (default 5)")] int nStrikes = 5,
            [ExcelArgument(Description = "Log-moneyness step (default 0.10 = 10%)")] double step = 0.10)
        {
            try
            {
                var evalQL  = QLHelper.ToQLDate(evalDate);
                var process = QLHelper.BuildBSMProcess(spot, dividendYield, riskFreeRate, 0.20, evalQL);
                int total   = 2 * nStrikes + 1;
                var result  = new object[total, 2];

                for (int i = 0; i < total; i++)
                {
                    double logM  = (i - nStrikes) * step;
                    double K     = spot * SM.Exp(logM);
                    result[i, 0] = K;
                    try
                    {
                        var opt = BuildHestonEuropean("call", spot, K, riskFreeRate, dividendYield,
                                                      v0, kappa, theta, sigma, rho, evalDate, maturityDate);
                        double price = opt.NPV();
                        var eu  = new VanillaOption(
                                      new PlainVanillaPayoff(Option.Type.Call, K),
                                      new EuropeanExercise(QLHelper.ToQLDate(maturityDate)));
                        result[i, 1] = eu.impliedVolatility(price, process, 1e-5, 300, 1e-5, 5.0);
                    }
                    catch { result[i, 1] = "N/A"; }
                }
                return result;
            }
            catch (Exception ex) { return new object[,] { { $"#QL_ERR: {ex.Message}" } }; }
        }
    }
}
