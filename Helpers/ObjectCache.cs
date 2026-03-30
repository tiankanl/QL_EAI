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
        private static readonly ConcurrentDictionary<string, YieldTermStructureHandle> _curves = new();

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
    }
}
