/*
 * AddInEvents.cs — ExcelDNA add-in lifecycle hooks.
 *
 * AutoOpen  : nothing special needed; ExcelDNA registers all [ExcelFunction] methods automatically.
 * AutoClose : dispose every cached QuantLib native handle eagerly so that the C++ destructors
 *             run all at once during unload rather than one-by-one through GC finalizers.
 *             Without this, closing Excel with a large curve cache causes a multi-second hang
 *             while the finalizer thread works through every YieldTermStructureHandle.
 */

using ExcelDna.Integration;
using QuantLibExcelAddin.Helpers;

namespace QuantLibExcelAddin
{
    public class AddInEvents : IExcelAddIn
    {
        public void AutoOpen()  { /* ExcelDNA handles function registration automatically. */ }

        public void AutoClose() => ObjectCache.ClearAll();
    }
}
