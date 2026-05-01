using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal static class ComObjectReleaseService
    {
        internal static void Release(object comObject)
        {
            ReleaseCore(comObject, finalRelease: false);
        }

        internal static void FinalRelease(object comObject)
        {
            ReleaseCore(comObject, finalRelease: true);
        }

        private static void ReleaseCore(object comObject, bool finalRelease)
        {
            if (comObject == null)
            {
                return;
            }

            try
            {
                if (!Marshal.IsComObject(comObject))
                {
                    return;
                }

                if (finalRelease)
                {
                    Marshal.FinalReleaseComObject(comObject);
                    return;
                }

                Marshal.ReleaseComObject(comObject);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("ComObjectReleaseService." + (finalRelease ? "FinalRelease" : "Release") + " failed: " + ex);
            }
        }
    }
}
