/*
 * MathFunctions.cs — Statistical distributions, interpolation, and Black formula utilities.
 * Mirrors ql/math/distributions/*, ql/math/interpolations/*, ql/pricingengines/blackformula.*
 *
 * NOTE: Distribution objects use .call(x) — the SWIG C# wrapper maps C++ operator() as call().
 *       Interpolation objects use QlArray (not DoubleVector) and .call(x, extrapolate).
 */

using System;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.Math
{
    public static class MathFunctions
    {
        // ─── Normal distribution ──────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_NormalCDF",
                       Category = "QuantLib — Math",
                       Description = "Cumulative standard normal distribution N(x). Highly accurate QuantLib implementation.")]
        public static object QL_NormalCDF(
            [ExcelArgument(Description = "Value x")] double x)
        {
            try { return new CumulativeNormalDistribution().call(x); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_NormalPDF",
                       Category = "QuantLib — Math",
                       Description = "Standard normal probability density function φ(x) = exp(-x²/2) / √(2π).")]
        public static object QL_NormalPDF(
            [ExcelArgument(Description = "Value x")] double x)
        {
            try { return new NormalDistribution().call(x); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_InverseNormalCDF",
                       Category = "QuantLib — Math",
                       Description = "Inverse cumulative standard normal N⁻¹(p). More accurate than Excel NORM.S.INV.")]
        public static object QL_InverseNormalCDF(
            [ExcelArgument(Description = "Probability p ∈ (0,1)")] double p)
        {
            try { return new InverseCumulativeNormal().call(p); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_StudentTCDF",
                       Category = "QuantLib — Math",
                       Description = "Cumulative Student-t distribution with ν degrees of freedom.")]
        public static object QL_StudentTCDF(
            [ExcelArgument(Description = "Value x")] double x,
            [ExcelArgument(Description = "Degrees of freedom ν (positive integer)")] int nu)
        {
            try { return new CumulativeStudentDistribution(nu).call(x); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_ChiSquaredCDF",
                       Category = "QuantLib — Math",
                       Description = "Cumulative chi-squared distribution with k degrees of freedom.")]
        public static object QL_ChiSquaredCDF(
            [ExcelArgument(Description = "Value x (≥ 0)")] double x,
            [ExcelArgument(Description = "Degrees of freedom k")] double k)
        {
            try { return new CumulativeChiSquareDistribution(k).call(x); }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BinomialCDF",
                       Category = "QuantLib — Math",
                       Description = "Cumulative binomial distribution P(X ≤ k) with success probability p and n trials.")]
        public static object QL_BinomialCDF(
            [ExcelArgument(Description = "Success probability p")] double p,
            [ExcelArgument(Description = "Number of trials n")] int n,
            [ExcelArgument(Description = "Observed successes k")] int k)
        {
            try
            {
                double cdf = 0.0;
                var dist = new BinomialDistribution(p, (uint)n);
                for (uint i = 0; i <= (uint)k; i++)
                    cdf += dist.call(i);
                return cdf;
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Interpolation ────────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_LinearInterpolate",
                       Category = "QuantLib — Math",
                       Description = "Linear interpolation at x from a set of (x, y) data points.\n" +
                                     "Provide xs and ys as matching Excel ranges (sorted ascending x).")]
        public static object QL_LinearInterpolate(
            [ExcelArgument(Description = "X values (sorted ascending)")] object xs,
            [ExcelArgument(Description = "Y values (matching x range)")] object ys,
            [ExcelArgument(Description = "Query point x₀")] double x0,
            [ExcelArgument(Description = "Allow extrapolation? (default false)")] bool extrapolate = false)
        {
            try
            {
                var xArr   = QLHelper.ToDoubleArray(xs);
                var yArr   = QLHelper.ToDoubleArray(ys);
                var interp = new LinearInterpolation(QLHelper.ToQlArray(xArr), QLHelper.ToQlArray(yArr));
                return interp.call(x0, extrapolate);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_CubicSplineInterpolate",
                       Category = "QuantLib — Math",
                       Description = "Natural cubic spline interpolation at x (C² continuity).\n" +
                                     "Provide xs and ys as matching Excel ranges.")]
        public static object QL_CubicSplineInterpolate(
            [ExcelArgument(Description = "X values (sorted ascending)")] object xs,
            [ExcelArgument(Description = "Y values (matching x range)")] object ys,
            [ExcelArgument(Description = "Query point x₀")] double x0,
            [ExcelArgument(Description = "Allow extrapolation? (default false)")] bool extrapolate = false)
        {
            try
            {
                var xArr   = QLHelper.ToDoubleArray(xs);
                var yArr   = QLHelper.ToDoubleArray(ys);
                var interp = new CubicNaturalSpline(QLHelper.ToQlArray(xArr), QLHelper.ToQlArray(yArr));
                return interp.call(x0, extrapolate);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_LogLinearInterpolate",
                       Category = "QuantLib — Math",
                       Description = "Log-linear interpolation at x (useful for discount factors; y values must be > 0).")]
        public static object QL_LogLinearInterpolate(
            [ExcelArgument(Description = "X values (sorted ascending)")] object xs,
            [ExcelArgument(Description = "Y values (must be > 0)")] object ys,
            [ExcelArgument(Description = "Query point x₀")] double x0,
            [ExcelArgument(Description = "Allow extrapolation? (default false)")] bool extrapolate = false)
        {
            try
            {
                var xArr   = QLHelper.ToDoubleArray(xs);
                var yArr   = QLHelper.ToDoubleArray(ys);
                var interp = new LogLinearInterpolation(QLHelper.ToQlArray(xArr), QLHelper.ToQlArray(yArr));
                return interp.call(x0, extrapolate);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BackwardFlatInterpolate",
                       Category = "QuantLib — Math",
                       Description = "Backward-flat (right-continuous step) interpolation.\n" +
                                     "Returns the y value of the node immediately to the right of x₀.")]
        public static object QL_BackwardFlatInterpolate(
            [ExcelArgument(Description = "X values")] object xs,
            [ExcelArgument(Description = "Y values")] object ys,
            [ExcelArgument(Description = "Query point x₀")] double x0,
            [ExcelArgument(Description = "Allow extrapolation? (default true)")] bool extrapolate = true)
        {
            try
            {
                var xArr   = QLHelper.ToDoubleArray(xs);
                var yArr   = QLHelper.ToDoubleArray(ys);
                var interp = new BackwardFlatInterpolation(QLHelper.ToQlArray(xArr), QLHelper.ToQlArray(yArr));
                return interp.call(x0, extrapolate);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_CubicSplineDerivative",
                       Category = "QuantLib — Math",
                       Description = "First derivative of a natural cubic spline interpolant at x₀.")]
        public static object QL_CubicSplineDerivative(
            [ExcelArgument(Description = "X values (sorted ascending)")] object xs,
            [ExcelArgument(Description = "Y values")] object ys,
            [ExcelArgument(Description = "Query point x₀")] double x0,
            [ExcelArgument(Description = "Allow extrapolation? (default false)")] bool extrapolate = false)
        {
            try
            {
                var interp = new CubicNaturalSpline(
                    QLHelper.ToQlArray(QLHelper.ToDoubleArray(xs)),
                    QLHelper.ToQlArray(QLHelper.ToDoubleArray(ys)));
                return interp.derivative(x0, extrapolate);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        // ─── Descriptive statistics ───────────────────────────────────────────────

        [ExcelFunction(Name = "QL_Statistics",
                       Category = "QuantLib — Math",
                       Description = "Compute descriptive statistics on a data range.\n" +
                                     "Stat: mean | stddev | variance | skewness | kurtosis | min | max.")]
        public static object QL_Statistics(
            [ExcelArgument(Description = "Data range (1-D)")] object data,
            [ExcelArgument(Description = "Statistic: mean, stddev, variance, skewness, kurtosis, min, max")] string stat)
        {
            try
            {
                var vals = QLHelper.ToDoubleArray(data);
                var s    = new IncrementalStatistics();
                foreach (var v in vals) s.add(v);
                return stat.Trim().ToLowerInvariant() switch
                {
                    "mean"      => s.mean(),
                    "stddev"    => s.standardDeviation(),
                    "variance"  => s.variance(),
                    "skewness"  => s.skewness(),
                    "kurtosis"  => s.kurtosis(),
                    "min"       => s.min(),
                    "max"       => s.max(),
                    _           => s.mean(),
                };
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }
    }
}
