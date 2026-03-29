/*
 * QLHelper.cs — shared conversion utilities for all QuantLib Excel UDF modules.
 *
 * All public helpers are internal so they are only accessible within
 * the QuantLibExcelAddin assembly and do not pollute the Excel function list.
 */

using System;
using QuantLib;

namespace QuantLibExcelAddin.Helpers
{
    internal static class QLHelper
    {
        // ─── Date helpers ────────────────────────────────────────────────────────

        /// <summary>Convert an Excel serial-date number to a QuantLib Date.</summary>
        internal static Date ToQLDate(double excelSerial)
        {
            DateTime dt = DateTime.FromOADate(excelSerial);
            return new Date(dt.Day, (Month)dt.Month, dt.Year);
        }

        /// <summary>Convert a QuantLib Date back to an Excel serial-date number.</summary>
        internal static double ToExcelDate(Date d) =>
            new DateTime(d.year(), (int)d.month(), d.dayOfMonth()).ToOADate();

        // ─── Option helpers ──────────────────────────────────────────────────────

        /// <summary>Parse "call"/"c" → Call; anything else → Put.</summary>
        internal static Option.Type ParseOptionType(string s) =>
            s.Trim().ToLowerInvariant().StartsWith("c")
                ? Option.Type.Call
                : Option.Type.Put;

        // ─── Barrier helpers ─────────────────────────────────────────────────────

        /// <summary>
        /// Parse barrier type string:
        ///   "DownIn" / "di"  → DownIn
        ///   "DownOut" / "do" → DownOut
        ///   "UpIn" / "ui"    → UpIn
        ///   "UpOut" / "uo"   → UpOut
        /// </summary>
        internal static QuantLib.Barrier.Type ParseBarrierType(string s)
        {
            return s.Trim().ToLowerInvariant() switch
            {
                "downin"   or "di" => QuantLib.Barrier.Type.DownIn,
                "downtown" or "do" => QuantLib.Barrier.Type.DownOut,
                "upin"     or "ui" => QuantLib.Barrier.Type.UpIn,
                "upout"    or "uo" => QuantLib.Barrier.Type.UpOut,
                _ => s.Contains("down", StringComparison.OrdinalIgnoreCase)
                        ? (s.Contains("in", StringComparison.OrdinalIgnoreCase)
                           ? QuantLib.Barrier.Type.DownIn : QuantLib.Barrier.Type.DownOut)
                        : (s.Contains("in", StringComparison.OrdinalIgnoreCase)
                           ? QuantLib.Barrier.Type.UpIn : QuantLib.Barrier.Type.UpOut)
            };
        }

        internal static DoubleBarrier.Type ParseDoubleBarrierType(string s)
        {
            return s.Trim().ToLowerInvariant() switch
            {
                "knockin"  or "ki" or "knock-in"  => DoubleBarrier.Type.KnockIn,
                "knockout" or "ko" or "knock-out" => DoubleBarrier.Type.KnockOut,
                "kiko"                            => DoubleBarrier.Type.KIKO,
                "koki"                            => DoubleBarrier.Type.KOKI,
                _                                 => DoubleBarrier.Type.KnockOut,
            };
        }

        // ─── Compounding / Frequency helpers ────────────────────────────────────

        /// <summary>
        /// Parse compounding convention string.
        ///   "Continuous" / "cont" / "c"
        ///   "Compounded" / "comp" / "annual" (default)
        ///   "Simple"
        ///   "SimpleThenCompounded" / "stc"
        /// </summary>
        internal static Compounding ParseCompounding(string s)
        {
            return s.Trim().ToLowerInvariant() switch
            {
                "continuous" or "cont" or "c" => Compounding.Continuous,
                "simple"                      => Compounding.Simple,
                "simplethencompounded" or "stc" => Compounding.SimpleThenCompounded,
                _                             => Compounding.Compounded,
            };
        }

        /// <summary>
        /// Parse frequency string:
        ///   "Annual" / "1" / "a"
        ///   "Semiannual" / "2" / "semi"
        ///   "Quarterly" / "4" / "q"
        ///   "Monthly" / "12" / "m"
        ///   "Daily" / "365"
        ///   "Once" (zero-coupon / no coupon)
        /// </summary>
        internal static Frequency ParseFrequency(string s)
        {
            return s.Trim().ToLowerInvariant() switch
            {
                "annual"     or "1"   or "a"    => Frequency.Annual,
                "semiannual" or "2"   or "semi"  => Frequency.Semiannual,
                "quarterly"  or "4"   or "q"    => Frequency.Quarterly,
                "monthly"    or "12"  or "m"    => Frequency.Monthly,
                "bimonthly"  or "6"              => Frequency.Bimonthly,
                "weekly"     or "52"  or "w"    => Frequency.Weekly,
                "daily"      or "365" or "d"    => Frequency.Daily,
                "once"       or "0"              => Frequency.Once,
                _                               => Frequency.Semiannual,
            };
        }

        // ─── Day-counter helper ──────────────────────────────────────────────────

        /// <summary>
        /// Parse day-counter string.
        ///   "Actual365" / "act365" / "a365"
        ///   "Actual360" / "act360" / "a360"
        ///   "Thirty360" / "30/360" / "30360"
        ///   "ActualActual" / "actact" / "aa" (uses ISMA convention)
        ///   "Business252" / "b252"
        /// Default: Actual365Fixed.
        /// </summary>
        internal static DayCounter ParseDayCounter(string s)
        {
            return s.Trim().ToLowerInvariant() switch
            {
                "actual360" or "act360" or "a360"         => (DayCounter)new Actual360(),
                "thirty360" or "30/360" or "30360"        => new Thirty360(Thirty360.Convention.BondBasis),
                "actualactual" or "actact" or "aa" or "isma" => new ActualActual(ActualActual.Convention.ISMA),
                "business252" or "b252"                   => new Business252(),
                _                                         => new Actual365Fixed(),
            };
        }

        // ─── Process / curve builders ────────────────────────────────────────────

        /// <summary>
        /// Build a Black-Scholes-Merton process from flat term structures.
        /// Sets the global QuantLib evaluation date as a side effect.
        /// </summary>
        internal static BlackScholesMertonProcess BuildBSMProcess(
            double spot, double dividendYield, double riskFreeRate, double volatility,
            Date evalDate)
        {
            Settings.instance().setEvaluationDate(evalDate);

            var calendar   = new TARGET();
            var dayCounter = new Actual365Fixed();

            var qHandle  = new QuoteHandle(new SimpleQuote(spot));
            var rtsHandle = new YieldTermStructureHandle(
                                new FlatForward(evalDate, riskFreeRate, dayCounter));
            var qtsHandle = new YieldTermStructureHandle(
                                new FlatForward(evalDate, dividendYield, dayCounter));
            var volHandle = new BlackVolTermStructureHandle(
                                new BlackConstantVol(evalDate, calendar, volatility, dayCounter));

            return new BlackScholesMertonProcess(qHandle, qtsHandle, rtsHandle, volHandle);
        }

        /// <summary>
        /// Build a flat <see cref="YieldTermStructureHandle"/> from a continuously-compounded rate.
        /// Sets the global QuantLib evaluation date as a side effect.
        /// </summary>
        internal static YieldTermStructureHandle FlatCurve(Date evalDate, double rate)
        {
            Settings.instance().setEvaluationDate(evalDate);
            return new YieldTermStructureHandle(
                       new FlatForward(evalDate, rate, new Actual365Fixed()));
        }

        // ─── Excel array helpers ─────────────────────────────────────────────────

        /// <summary>Convert a 1-D Excel range (object[] or double[]) to double[].</summary>
        internal static double[] ToDoubleArray(object input)
        {
            if (input is double d)   return new[] { d };
            if (input is double[] a) return a;

            if (input is object[] oa)
            {
                var result = new double[oa.Length];
                for (int i = 0; i < oa.Length; i++)
                    result[i] = Convert.ToDouble(oa[i]);
                return result;
            }
            if (input is object[,] oa2)
            {
                int rows = oa2.GetLength(0), cols = oa2.GetLength(1);
                var result = new double[rows * cols];
                int k = 0;
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        result[k++] = Convert.ToDouble(oa2[r, c]);
                return result;
            }
            throw new ArgumentException($"Cannot convert {input?.GetType().Name} to double[]");
        }

        /// <summary>Convert a 1-D Excel date range to QuantLib DateVector.</summary>
        internal static DateVector ToDateVector(object input)
        {
            var doubles = ToDoubleArray(input);
            var dv      = new DateVector();
            foreach (var v in doubles)
                dv.Add(ToQLDate(v));
            return dv;
        }

        /// <summary>Convert a double[] to a QuantLib DoubleVector.</summary>
        internal static DoubleVector ToDoubleVector(double[] arr)
        {
            var dv = new DoubleVector();
            foreach (var v in arr) dv.Add(v);
            return dv;
        }

        /// <summary>Convert a double[] to a QuantLib QlArray (used by interpolation classes).</summary>
        internal static QlArray ToQlArray(double[] arr)
        {
            var qa = new QlArray((uint)arr.Length);
            for (uint i = 0; i < arr.Length; i++)
                qa.set(i, arr[(int)i]);
            return qa;
        }
    }
}
