/*
 * EuropeanOptions.cs — European vanilla options, all BSM Greeks, and implied volatility.
 * Mirrors ql/instruments/vanillaoption.* and ql/pricingengines/vanilla/analyticeuropeanengine.*
 */

using System;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.Options
{
    public static class EuropeanOptions
    {
        // ─── Internal builder ─────────────────────────────────────────────────────

        /// <summary>Price a European vanilla option and return the requested Greek or NPV.</summary>
        private static double BSMGreek(
            string optionType, double spot, double strike,
            double riskFreeRate, double dividendYield, double volatility,
            double evalSerial, double maturitySerial, string greek)
        {
            var evalDate = QLHelper.ToQLDate(evalSerial);
            var matDate  = QLHelper.ToQLDate(maturitySerial);
            var process  = QLHelper.BuildBSMProcess(spot, dividendYield, riskFreeRate, volatility, evalDate);
            var payoff   = new PlainVanillaPayoff(QLHelper.ParseOptionType(optionType), strike);
            var exercise = new EuropeanExercise(matDate);
            var option   = new VanillaOption(payoff, exercise);
            option.setPricingEngine(new AnalyticEuropeanEngine(process));

            return greek switch
            {
                "delta"       => option.delta(),
                "gamma"       => option.gamma(),
                "vega"        => option.vega(),
                "theta"       => option.theta(),
                "rho"         => option.rho(),
                "elasticity"  => option.elasticity(),
                "thetaperday" => option.thetaPerDay(),
                _             => option.NPV(),
            };
        }

        // ─── Price ────────────────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_BSMPrice",
                       Category = "QuantLib — European Options",
                       Description = "Price a European vanilla option using the Black-Scholes-Merton analytic engine.")]
        public static object QL_BSMPrice(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Underlying spot price")] double spot,
            [ExcelArgument(Description = "Strike price")] double strike,
            [ExcelArgument(Description = "Risk-free rate (decimal, e.g. 0.05)")] double riskFreeRate,
            [ExcelArgument(Description = "Continuous dividend yield (decimal)")] double dividendYield,
            [ExcelArgument(Description = "Volatility (decimal, e.g. 0.20)")] double volatility,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Maturity / expiry date (Excel date)")] double maturityDate)
        {
            try { return BSMGreek(optionType, spot, strike, riskFreeRate, dividendYield, volatility, evalDate, maturityDate, "npv"); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Greeks ───────────────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_BSMDelta",
                       Category = "QuantLib — European Options",
                       Description = "Delta (∂V/∂S) of a European option via BSM.")]
        public static object QL_BSMDelta(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate)
        {
            try { return BSMGreek(optionType, spot, strike, riskFreeRate, dividendYield, volatility, evalDate, maturityDate, "delta"); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BSMGamma",
                       Category = "QuantLib — European Options",
                       Description = "Gamma (∂²V/∂S²) of a European option via BSM.")]
        public static object QL_BSMGamma(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate)
        {
            try { return BSMGreek(optionType, spot, strike, riskFreeRate, dividendYield, volatility, evalDate, maturityDate, "gamma"); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BSMVega",
                       Category = "QuantLib — European Options",
                       Description = "Vega (∂V/∂σ) of a European option via BSM.")]
        public static object QL_BSMVega(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate)
        {
            try { return BSMGreek(optionType, spot, strike, riskFreeRate, dividendYield, volatility, evalDate, maturityDate, "vega"); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BSMTheta",
                       Category = "QuantLib — European Options",
                       Description = "Theta (∂V/∂t) of a European option via BSM (per year).")]
        public static object QL_BSMTheta(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate)
        {
            try { return BSMGreek(optionType, spot, strike, riskFreeRate, dividendYield, volatility, evalDate, maturityDate, "theta"); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BSMThetaPerDay",
                       Category = "QuantLib — European Options",
                       Description = "Theta per calendar day (= Theta/365) via BSM.")]
        public static object QL_BSMThetaPerDay(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate)
        {
            try { return BSMGreek(optionType, spot, strike, riskFreeRate, dividendYield, volatility, evalDate, maturityDate, "thetaperday"); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BSMRho",
                       Category = "QuantLib — European Options",
                       Description = "Rho (∂V/∂r) of a European option via BSM.")]
        public static object QL_BSMRho(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate)
        {
            try { return BSMGreek(optionType, spot, strike, riskFreeRate, dividendYield, volatility, evalDate, maturityDate, "rho"); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BSMVanna",
                       Category = "QuantLib — European Options",
                       Description = "Vanna (∂²V/∂S∂σ) of a European option via BSM.")]
        public static object QL_BSMVanna(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate)
        {
            try { return BSMGreek(optionType, spot, strike, riskFreeRate, dividendYield, volatility, evalDate, maturityDate, "vanna"); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BSMVolga",
                       Category = "QuantLib — European Options",
                       Description = "Volga / Vomma (∂²V/∂σ²) of a European option via BSM.")]
        public static object QL_BSMVolga(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate)
        {
            try { return BSMGreek(optionType, spot, strike, riskFreeRate, dividendYield, volatility, evalDate, maturityDate, "volga"); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BSMElasticity",
                       Category = "QuantLib — European Options",
                       Description = "Elasticity (Λ = Δ × S/V) of a European option via BSM.")]
        public static object QL_BSMElasticity(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate)
        {
            try { return BSMGreek(optionType, spot, strike, riskFreeRate, dividendYield, volatility, evalDate, maturityDate, "elasticity"); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BSMStrikeSensitivity",
                       Category = "QuantLib — European Options",
                       Description = "Strike sensitivity (∂V/∂K) of a European option via BSM.")]
        public static object QL_BSMStrikeSensitivity(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike")] double strike,
            [ExcelArgument(Description = "Risk-free rate")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield")] double dividendYield,
            [ExcelArgument(Description = "Volatility")] double volatility,
            [ExcelArgument(Description = "Evaluation date")] double evalDate,
            [ExcelArgument(Description = "Maturity date")] double maturityDate)
        {
            try { return BSMGreek(optionType, spot, strike, riskFreeRate, dividendYield, volatility, evalDate, maturityDate, "strikesens"); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Implied volatility ───────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_BSMImpliedVol",
                       Category = "QuantLib — European Options",
                       Description = "Implied Black-Scholes volatility from a market option price using Newton-Raphson root finding.")]
        public static object QL_BSMImpliedVol(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Market price of the option")] double marketPrice,
            [ExcelArgument(Description = "Underlying spot price")] double spot,
            [ExcelArgument(Description = "Strike price")] double strike,
            [ExcelArgument(Description = "Risk-free rate (decimal)")] double riskFreeRate,
            [ExcelArgument(Description = "Dividend yield (decimal)")] double dividendYield,
            [ExcelArgument(Description = "Evaluation date (Excel date)")] double evalDate,
            [ExcelArgument(Description = "Maturity date (Excel date)")] double maturityDate,
            [ExcelArgument(Description = "Accuracy (optional, default 1e-6)")] double accuracy = 1e-6,
            [ExcelArgument(Description = "Max evaluations (optional, default 200)")] int maxEvaluations = 200,
            [ExcelArgument(Description = "Min vol guess (optional, default 1e-4)")] double minVol = 1e-4,
            [ExcelArgument(Description = "Max vol guess (optional, default 4.0)")] double maxVol = 4.0)
        {
            try
            {
                var evalQL = QLHelper.ToQLDate(evalDate);
                var matQL  = QLHelper.ToQLDate(maturityDate);
                var process = QLHelper.BuildBSMProcess(spot, dividendYield, riskFreeRate, 0.20, evalQL);
                var payoff  = new PlainVanillaPayoff(QLHelper.ParseOptionType(optionType), strike);
                var option  = new VanillaOption(payoff, new EuropeanExercise(matQL));
                return option.impliedVolatility(marketPrice, process, accuracy, (uint)maxEvaluations, minVol, maxVol);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── European digital options ─────────────────────────────────────────────

        [ExcelFunction(Name = "QL_EuropeanDigitalCash",
                       Category = "QuantLib — European Options",
                       Description = "Price a European cash-or-nothing digital option via BSM analytic engine.\n" +
                                     "Pays 'cashPayoff' if the underlying finishes in-the-money at expiry.")]
        public static object QL_EuropeanDigitalCash(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike / barrier level")] double strike,
            [ExcelArgument(Description = "Cash payoff amount")] double cashPayoff,
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
                var option  = new VanillaOption(payoff, new EuropeanExercise(QLHelper.ToQLDate(maturityDate)));
                option.setPricingEngine(new AnalyticEuropeanEngine(process));
                return option.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_EuropeanDigitalAsset",
                       Category = "QuantLib — European Options",
                       Description = "Price a European asset-or-nothing digital option via BSM analytic engine.\n" +
                                     "Pays the asset value if the underlying finishes in-the-money at expiry.")]
        public static object QL_EuropeanDigitalAsset(
            [ExcelArgument(Description = "\"call\" or \"put\"")] string optionType,
            [ExcelArgument(Description = "Spot")] double spot,
            [ExcelArgument(Description = "Strike / barrier level")] double strike,
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
                var payoff  = new AssetOrNothingPayoff(QLHelper.ParseOptionType(optionType), strike);
                var option  = new VanillaOption(payoff, new EuropeanExercise(QLHelper.ToQLDate(maturityDate)));
                option.setPricingEngine(new AnalyticEuropeanEngine(process));
                return option.NPV();
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }
    }
}
