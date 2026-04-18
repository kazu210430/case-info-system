using System;
using System.Runtime.InteropServices;
using CaseInfoSystem.ExcelAddIn.Infrastructure;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    [ComVisible(true)]
    [Guid("A8A8D3B9-3396-47C7-A3D0-949B5C4A4811")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IKernelAddInAutomation
    {
        void ShowKernelHomeFromSheet();
        void ReflectKernelUserDataToAccountingSet();
        void ReflectKernelUserDataToBaseHome();
    }

    /// <summary>
    [ComVisible(true)]
    [Guid("5E92E4D5-B8AB-49D8-BF09-E611857D4F07")]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class KernelAutomationService : IKernelAddInAutomation
    {
        private readonly ThisAddIn _addIn;

        internal KernelAutomationService(ThisAddIn addIn)
        {
            _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
        }

        public void ShowKernelHomeFromSheet()
        {
            try
            {
                _addIn.ShowKernelHomeFromAutomation();
            }
            catch (Exception ex)
            {
                _addIn.LogAutomationFailure("Kernel automation failed. method=ShowKernelHomeFromSheet", ex);
                throw;
            }
        }

        public void ReflectKernelUserDataToAccountingSet()
        {
            try
            {
                _addIn.ReflectKernelUserDataToAccountingSet();
            }
            catch (Exception ex)
            {
                _addIn.LogAutomationFailure("Kernel automation failed. method=ReflectKernelUserDataToAccountingSet", ex);
                throw;
            }
        }

        public void ReflectKernelUserDataToBaseHome()
        {
            try
            {
                _addIn.ReflectKernelUserDataToBaseHome();
            }
            catch (Exception ex)
            {
                _addIn.LogAutomationFailure("Kernel automation failed. method=ReflectKernelUserDataToBaseHome", ex);
                throw;
            }
        }
    }
}
