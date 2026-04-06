/*
 * ObjectCache.cs — Static in-process cache for QuantLib yield curve handles.
 *
 * Pattern:
 *   1. QL_BuildSwapCurve / QL_BuildFlatCurve bootstraps a curve and stores it here,
 *      returning a deterministic string key derived from the inputs.
 *   2. Pricing functions (QL_FixedBondCashflowsH, QL_FRNPriceH, etc.) look up the
 *      key and get back a ready-to-use YieldTermStructureHandle — no re-bootstrapping.
 *   3. Because the key is deterministic, unchanged inputs return the same key and
 *      Excel's dependency graph does not trigger unnecessary recalculations downstream.
 */

using System.Collections.Concurrent;
using QuantLib;

namespace QuantLibExcelAddin.Helpers
{
    internal static class ObjectCache
    {
        private static readonly ConcurrentDictionary<string, YieldTermStructureHandle> _curves     = new();
        private static readonly ConcurrentDictionary<string, double[]>                 _nodeDates  = new();

        /// <summary>Store a curve handle under <paramref name="key"/>; returns the key.</summary>
        internal static string StoreCurve(string key, YieldTermStructureHandle handle)
        {
            _curves[key] = handle;
            return key;
        }

        /// <summary>
        /// Retrieve a previously stored curve handle.
        /// Throws a descriptive ArgumentException if the key is not found
        /// (e.g. after an Excel session restart — just recalculate the builder cell).
        /// </summary>
        internal static YieldTermStructureHandle GetCurve(string key)
        {
            if (_curves.TryGetValue(key, out var handle))
                return handle;
            throw new ArgumentException(
                $"Curve handle '{key}' not found in cache. " +
                 "Force-recalculate the QL_BuildSwapCurve / QL_BuildFlatCurve cell (Ctrl+Alt+F9).");
        }

        internal static bool HasCurve(string key) => _curves.ContainsKey(key);

        // ─── Node-date storage ────────────────────────────────────────────────────

        /// <summary>
        /// Store the pillar dates (as Excel serials) for the curve identified by <paramref name="key"/>.
        /// Called by each curve builder so that QL_CurveNodes can retrieve them later.
        /// </summary>
        internal static void StoreNodeDates(string key, double[] excelDates)
            => _nodeDates[key] = excelDates;

        /// <summary>
        /// Retrieve the stored pillar dates for a curve, or <c>null</c> if none were stored
        /// (e.g. the curve was built in an older session before this feature existed).
        /// </summary>
        internal static double[]? GetNodeDates(string key)
            => _nodeDates.TryGetValue(key, out var dates) ? dates : null;

        // ─── Shutdown ─────────────────────────────────────────────────────────────

        /// <summary>
        /// Dispose every cached native handle and clear both dictionaries.
        /// Call this from IExcelAddIn.AutoClose() so that QuantLib's C++ destructors
        /// run eagerly on unload rather than one-by-one through GC finalizers,
        /// which is what makes Excel slow to close when the cache is large.
        /// </summary>
        internal static void ClearAll()
        {
            foreach (var handle in _curves.Values)
            {
                try { handle.Dispose(); }
                catch { /* never let a bad handle block shutdown */ }
            }
            _curves.Clear();
            _nodeDates.Clear();
        }
    }
}
