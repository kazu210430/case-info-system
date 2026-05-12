using System;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.ExcelAddIn.UI;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelHomeFormHost
    {
        private readonly KernelWorkbookService _kernelWorkbookService;
        private readonly KernelCaseCreationCommandService _kernelCaseCreationCommandService;
        private readonly Logger _logger;
        private KernelHomeForm _form;

        internal KernelHomeFormHost(
            KernelWorkbookService kernelWorkbookService,
            KernelCaseCreationCommandService kernelCaseCreationCommandService,
            Logger logger)
        {
            _kernelWorkbookService = kernelWorkbookService ?? throw new ArgumentNullException(nameof(kernelWorkbookService));
            _kernelCaseCreationCommandService = kernelCaseCreationCommandService ?? throw new ArgumentNullException(nameof(kernelCaseCreationCommandService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        internal KernelHomeForm Current
        {
            get { return _form; }
        }

        internal KernelHomeForm GetOrCreate(bool clearBindingOnNewSession)
        {
            DisposeHiddenFormBeforeRecreation();

            if (_form == null || _form.IsDisposed)
            {
                if (clearBindingOnNewSession)
                {
                    _kernelWorkbookService.ClearHomeWorkbookBinding("ThisAddIn.ShowKernelHomePlaceholder.NewSession");
                }

                _form = new KernelHomeForm(_kernelWorkbookService, _kernelCaseCreationCommandService, _logger);
            }

            return _form;
        }

        internal void ReloadCurrent()
        {
            if (_form == null || _form.IsDisposed)
            {
                return;
            }

            _form.ReloadSettings();
            _form.Invalidate(true);
            _form.Update();
        }

        internal void ShowAndActivateCurrent()
        {
            if (_form == null || _form.IsDisposed)
            {
                return;
            }

            if (!_form.Visible)
            {
                _form.Show();
            }

            _form.Activate();
            _form.BringToFront();
        }

        internal void HideCurrent()
        {
            if (_form == null || _form.IsDisposed || !_form.Visible)
            {
                return;
            }

            try
            {
                _form.Hide();
            }
            catch (Exception ex)
            {
                _logger.Error("HideKernelHomePlaceholder failed.", ex);
            }
        }

        internal void CloseOnShutdown()
        {
            if (_form != null && !_form.IsDisposed)
            {
                _form.Close();
                _form = null;
            }
        }

        private void DisposeHiddenFormBeforeRecreation()
        {
            if (_form == null || _form.IsDisposed || _form.Visible)
            {
                return;
            }

            try
            {
                _form.PrepareForSilentDispose();
                _form.Dispose();
            }
            catch (Exception ex)
            {
                _logger.Error("KernelHomeForm dispose before recreation failed.", ex);
            }
            finally
            {
                _form = null;
            }
        }
    }
}
