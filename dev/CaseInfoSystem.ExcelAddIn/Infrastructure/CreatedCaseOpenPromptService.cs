using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.UI;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class CreatedCaseOpenPromptService
	{
		private const int PromptDelayAfterFolderOpenMs = 0;

		private readonly Logger _logger;

		private Form _temporarilyMinimizedOwner;

		private FormWindowState _temporarilyMinimizedOwnerState = FormWindowState.Normal;

		internal CreatedCaseOpenPromptService (Logger logger)
		{
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal CreatedCaseOpenDecision ConfirmOpenCreatedCase (Stopwatch commandStopwatch)
		{
			Thread.Sleep (0);
			PreparePromptOwnerForModalDisplay ();
			_logger.Info ("Created CASE open prompt showing requested. stage=before-showdialog, elapsedMs=" + GetElapsedMilliseconds (commandStopwatch));
			try {
				CreatedCaseOpenDecision createdCaseOpenDecision = CreatedCaseOpenPromptForm.ShowPrompt (null, OnPromptShown (commandStopwatch));
				if (createdCaseOpenDecision == CreatedCaseOpenDecision.Skip) {
					ClearTrackedPromptOwner ();
				}
				_logger.Info ("Created CASE open prompt completed. decision=" + createdCaseOpenDecision);
				return createdCaseOpenDecision;
			} catch {
				RestorePromptOwnerIfNeeded ();
				throw;
			}
		}

		private static IWin32Window ResolvePromptOwner ()
		{
			FormCollection openForms = Application.OpenForms;
			if (openForms == null || openForms.Count == 0) {
				return null;
			}
			for (int num = openForms.Count - 1; num >= 0; num--) {
				Form form = openForms [num];
				if (form != null && !form.IsDisposed && form.Visible) {
					return form;
				}
			}
			return null;
		}

		private void PreparePromptOwnerForModalDisplay ()
		{
			ClearTrackedPromptOwner ();
			if (ResolvePromptOwner () is Form form && !form.IsDisposed && form.Visible) {
				_temporarilyMinimizedOwner = form;
				_temporarilyMinimizedOwnerState = form.WindowState;
				form.WindowState = FormWindowState.Minimized;
			}
		}

		internal void RestorePromptOwnerIfNeeded ()
		{
			if (_temporarilyMinimizedOwner == null || _temporarilyMinimizedOwner.IsDisposed) {
				ClearTrackedPromptOwner ();
				return;
			}
			_temporarilyMinimizedOwner.WindowState = _temporarilyMinimizedOwnerState;
			_temporarilyMinimizedOwner.Activate ();
			ClearTrackedPromptOwner ();
		}

		private void ClearTrackedPromptOwner ()
		{
			_temporarilyMinimizedOwner = null;
			_temporarilyMinimizedOwnerState = FormWindowState.Normal;
		}

		private Action OnPromptShown (Stopwatch commandStopwatch)
		{
			return delegate {
				_logger.Info ("Created CASE open prompt shown. stage=shown, elapsedMs=" + GetElapsedMilliseconds (commandStopwatch));
			};
		}

		private static long GetElapsedMilliseconds (Stopwatch stopwatch)
		{
			return (stopwatch == null) ? 0L : stopwatch.ElapsedMilliseconds;
		}
	}
}
