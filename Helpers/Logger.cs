/*
 * Logger.cs — Lightweight file logger for the QuantLib Excel add-in.
 *
 * Usage (from any function):
 *   Logger.Info ("QL_BuildSwapCurve", $"Building curve key={key}");
 *   Logger.Warn ("ObjectCache",       "Cache miss — rebuilding");
 *   Logger.Error("QL_FRNPrice",       ex.Message);
 *
 * Logging is off by default (zero overhead).  Enable it from Excel:
 *   =QL_LogEnable("C:\Temp\ql_addin.log")
 *   =QL_LogDisable()
 *   =QL_LogClear()
 *   =QL_LogStatus()
 *
 * Thread safety: a single lock serialises all writes so concurrent
 * Excel function calls don't interleave log lines.
 */

using System;
using System.IO;

namespace QuantLibExcelAddin.Helpers
{
    internal static class Logger
    {
        private static readonly object _lock = new();
        private static string?         _path;          // null = disabled

        // ─── State ────────────────────────────────────────────────────────────────

        internal static bool    IsEnabled => _path != null;
        internal static string? FilePath  => _path;

        // ─── Control ──────────────────────────────────────────────────────────────

        internal static void Enable(string path)
        {
            lock (_lock) { _path = path; }
            Info("Logger", $"Logging started → {path}");
        }

        internal static void Disable()
        {
            Info("Logger", "Logging stopped.");
            lock (_lock) { _path = null; }
        }

        /// <summary>Delete all content from the log file (keeps the file at the same path).</summary>
        internal static void Clear()
        {
            lock (_lock)
            {
                if (_path != null)
                    File.WriteAllText(_path, string.Empty);
            }
        }

        // ─── Write helpers ────────────────────────────────────────────────────────

        internal static void Info (string source, string message) => Write("INFO ", source, message);
        internal static void Warn (string source, string message) => Write("WARN ", source, message);
        internal static void Error(string source, string message) => Write("ERROR", source, message);

        private static void Write(string level, string source, string message)
        {
            if (_path == null) return;          // fast path — logging disabled

            var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}  [{level}]  [{source}]  {message}";
            lock (_lock)
            {
                try
                {
                    File.AppendAllText(_path, line + Environment.NewLine);
                }
                catch
                {
                    // Never let a logging failure crash a pricing function.
                }
            }
        }
    }
}
