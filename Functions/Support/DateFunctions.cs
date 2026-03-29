/*
 * DateFunctions.cs — Excel UDFs for date arithmetic, calendars, and day-count conventions.
 * Mirrors ql/time/ in the QuantLib source tree.
 */

using System;
using ExcelDna.Integration;
using QuantLib;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin.Functions.Support
{
    public static class DateFunctions
    {
        // ─── Calendar helpers ─────────────────────────────────────────────────────

        private static Calendar GetCalendar(string name) =>
            name.Trim().ToLowerInvariant() switch
            {
                "target" or "eur"               => (Calendar)new TARGET(),
                "us" or "usd" or "unitedstates" => new UnitedStates(UnitedStates.Market.Settlement),
                "uk" or "gbp" or "unitedkingdom" => new UnitedKingdom(),
                "ny" or "nyse"                  => new UnitedStates(UnitedStates.Market.NYSE),
                "uk/exchange" or "lse"          => new UnitedKingdom(UnitedKingdom.Market.Exchange),
                "japan" or "jpy" or "tse"       => new Japan(),
                "china"                         => new China(),
                "australia"                     => new Australia(),
                "canada"                        => new Canada(),
                "germany"                       => new Germany(Germany.Market.FrankfurtStockExchange),
                "weekendsonly"                  => new WeekendsOnly(),
                "null" or "none"                => new NullCalendar(),
                _                               => new TARGET(),
            };

        private static BusinessDayConvention GetBDC(string s) =>
            s.Trim().ToLowerInvariant() switch
            {
                "following"          or "f"   => BusinessDayConvention.Following,
                "preceding"          or "p"   => BusinessDayConvention.Preceding,
                "modifiedfollowing"  or "mf"  => BusinessDayConvention.ModifiedFollowing,
                "modifiedpreceding"  or "mp"  => BusinessDayConvention.ModifiedPreceding,
                "unadjusted"         or "u"   => BusinessDayConvention.Unadjusted,
                _                             => BusinessDayConvention.ModifiedFollowing,
            };

        // ─── UDFs ─────────────────────────────────────────────────────────────────

        [ExcelFunction(Name = "QL_DateAdd",
                       Category = "QuantLib — Dates",
                       Description = "Advance a date by a tenor using a QuantLib calendar.\n" +
                                     "Tenor examples: '3M', '1Y', '6W', '90D'.\n" +
                                     "Calendar: TARGET (default), US, UK, Japan, etc.\n" +
                                     "Convention: ModifiedFollowing (default), Following, Preceding, Unadjusted.")]
        public static object QL_DateAdd(
            [ExcelArgument(Description = "Start date (Excel date number)")] double startDate,
            [ExcelArgument(Description = "Tenor string, e.g. '3M', '1Y', '6W'")] string tenor,
            [ExcelArgument(Description = "Calendar (optional, default TARGET)")] string calendar = "TARGET",
            [ExcelArgument(Description = "Business day convention (optional, default ModifiedFollowing)")] string convention = "ModifiedFollowing")
        {
            try
            {
                var d  = QLHelper.ToQLDate(startDate);
                var p  = new Period(tenor);
                var cal = GetCalendar(calendar);
                var bdc = GetBDC(convention);
                var result = cal.advance(d, p, bdc);
                return QLHelper.ToExcelDate(result);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_DaysBetween",
                       Category = "QuantLib — Dates",
                       Description = "Number of calendar days between two dates (d2 - d1).")]
        public static object QL_DaysBetween(
            [ExcelArgument(Description = "Start date")] double date1,
            [ExcelArgument(Description = "End date")]   double date2)
        {
            try
            {
                var d1 = QLHelper.ToQLDate(date1);
                var d2 = QLHelper.ToQLDate(date2);
                return (int)(d2 - d1);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_DayCountFraction",
                       Category = "QuantLib — Dates",
                       Description = "Year fraction between two dates using a QuantLib day-count convention.\n" +
                                     "Conventions: Actual365 (default), Actual360, Thirty360, ActualActual, Business252.")]
        public static object QL_DayCountFraction(
            [ExcelArgument(Description = "Start date")] double date1,
            [ExcelArgument(Description = "End date")]   double date2,
            [ExcelArgument(Description = "Day-count convention (optional, default Actual365)")] string dayCounter = "Actual365")
        {
            try
            {
                var d1 = QLHelper.ToQLDate(date1);
                var d2 = QLHelper.ToQLDate(date2);
                var dc = QLHelper.ParseDayCounter(dayCounter);
                return dc.yearFraction(d1, d2);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_IsBusinessDay",
                       Category = "QuantLib — Dates",
                       Description = "Returns TRUE if the date is a business day in the given calendar.")]
        public static object QL_IsBusinessDay(
            [ExcelArgument(Description = "Date to test")] double date,
            [ExcelArgument(Description = "Calendar (optional, default TARGET)")] string calendar = "TARGET")
        {
            try
            {
                var d   = QLHelper.ToQLDate(date);
                var cal = GetCalendar(calendar);
                return cal.isBusinessDay(d);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_NextBusinessDay",
                       Category = "QuantLib — Dates",
                       Description = "Returns the next business day on or after the given date.")]
        public static object QL_NextBusinessDay(
            [ExcelArgument(Description = "Reference date")] double date,
            [ExcelArgument(Description = "Calendar (optional, default TARGET)")] string calendar = "TARGET")
        {
            try
            {
                var d   = QLHelper.ToQLDate(date);
                var cal = GetCalendar(calendar);
                return QLHelper.ToExcelDate(cal.adjust(d, BusinessDayConvention.Following));
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_PreviousBusinessDay",
                       Category = "QuantLib — Dates",
                       Description = "Returns the previous business day on or before the given date.")]
        public static object QL_PreviousBusinessDay(
            [ExcelArgument(Description = "Reference date")] double date,
            [ExcelArgument(Description = "Calendar (optional, default TARGET)")] string calendar = "TARGET")
        {
            try
            {
                var d   = QLHelper.ToQLDate(date);
                var cal = GetCalendar(calendar);
                return QLHelper.ToExcelDate(cal.adjust(d, BusinessDayConvention.Preceding));
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_EndOfMonth",
                       Category = "QuantLib — Dates",
                       Description = "Returns the last business day of the month for the given date and calendar.")]
        public static object QL_EndOfMonth(
            [ExcelArgument(Description = "Reference date")] double date,
            [ExcelArgument(Description = "Calendar (optional, default TARGET)")] string calendar = "TARGET")
        {
            try
            {
                var d   = QLHelper.ToQLDate(date);
                var cal = GetCalendar(calendar);
                return QLHelper.ToExcelDate(cal.endOfMonth(d));
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_BusinessDaysBetween",
                       Category = "QuantLib — Dates",
                       Description = "Count business days between two dates in the given calendar.")]
        public static object QL_BusinessDaysBetween(
            [ExcelArgument(Description = "Start date")] double date1,
            [ExcelArgument(Description = "End date")]   double date2,
            [ExcelArgument(Description = "Calendar (optional, default TARGET)")] string calendar = "TARGET",
            [ExcelArgument(Description = "Include start date? (default false)")] bool includeFirst = false,
            [ExcelArgument(Description = "Include end date? (default true)")] bool includeLast = true)
        {
            try
            {
                var d1  = QLHelper.ToQLDate(date1);
                var d2  = QLHelper.ToQLDate(date2);
                var cal = GetCalendar(calendar);
                return cal.businessDaysBetween(d1, d2, includeFirst, includeLast);
            }
            catch (Exception ex) { return $"#QL_ERR: {ex.Message}"; }
        }

        [ExcelFunction(Name = "QL_Version",
                       Category = "QuantLib — Dates",
                       Description = "Returns the QuantLib SWIG wrapper version string.")]
        public static string QL_Version() =>
            "QuantLib-SWIG 4.4.1 (NQuantLibc.dll / ExcelDNA 1.8.0)";
    }
}
