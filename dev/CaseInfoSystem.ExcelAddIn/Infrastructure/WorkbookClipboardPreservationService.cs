using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class WorkbookClipboardPreservationService
	{
		private const int ClipboardRetryCount = 3;

		private const int ClipboardRetryDelayMs = 40;

		private const long MaxPreservedCellCount = 50000L;

		private readonly WorkbookRoleResolver _workbookRoleResolver;

		private readonly Logger _logger;

		internal WorkbookClipboardPreservationService (WorkbookRoleResolver workbookRoleResolver, Logger logger)
		{
			_workbookRoleResolver = workbookRoleResolver ?? throw new ArgumentNullException ("workbookRoleResolver");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal void PreserveCopiedValuesForClosingWorkbook (Workbook workbook)
		{
			if (!ShouldHandleWorkbook (workbook)) {
				return;
			}
			Range range = null;
			try {
				Microsoft.Office.Interop.Excel.Application application = workbook.Application;
				if (application != null && application.CutCopyMode == XlCutCopyMode.xlCopy) {
					range = application.Selection as Range;
					if (IsSupportedSelection (range, workbook, out var cellCount)) {
						string clipboardText = BuildClipboardText (range);
						WriteClipboardTextWithRetry (clipboardText);
						_logger.Info ("Workbook copy values preserved to clipboard before close. workbook=" + GetWorkbookKey (workbook) + ", cells=" + cellCount.ToString (CultureInfo.InvariantCulture));
					}
				}
			} catch (Exception exception) {
				_logger.Error ("Workbook clipboard preservation failed.", exception);
			} finally {
				ReleaseComObject (range);
			}
		}

		private bool ShouldHandleWorkbook (Workbook workbook)
		{
			return _workbookRoleResolver.IsCaseWorkbook (workbook) || _workbookRoleResolver.IsAccountingWorkbook (workbook);
		}

		private bool IsSupportedSelection (Range selectionRange, Workbook workbook, out long cellCount)
		{
			cellCount = 0L;
			if (selectionRange == null || workbook == null) {
				return false;
			}
			Worksheet worksheet = null;
			Workbook workbook2 = null;
			object obj = null;
			try {
				worksheet = selectionRange.Worksheet;
				workbook2 = ((worksheet == null) ? null : (worksheet.Parent as Workbook));
				if (workbook2 != workbook) {
					return false;
				}
				obj = selectionRange.Areas;
				dynamic val = obj;
				int num = Convert.ToInt32 (val.Count, CultureInfo.InvariantCulture);
				if (num != 1) {
					_logger.Info ("Workbook clipboard preservation skipped because multi-area selection is not supported. workbook=" + GetWorkbookKey (workbook));
					return false;
				}
				cellCount = Convert.ToInt64 ((dynamic)selectionRange.CountLarge, CultureInfo.InvariantCulture);
				if (cellCount < 1) {
					return false;
				}
				if (cellCount > 50000) {
					_logger.Info ("Workbook clipboard preservation skipped because selection is too large. workbook=" + GetWorkbookKey (workbook) + ", cells=" + cellCount.ToString (CultureInfo.InvariantCulture));
					return false;
				}
				return true;
			} finally {
				ReleaseComObject (obj);
				ReleaseComObject (workbook2);
				ReleaseComObject (worksheet);
			}
		}

		private static string BuildClipboardText (Range selectionRange)
		{
			object value = selectionRange.Value2;
			if (!(value is object[,] array)) {
				return SanitizeClipboardCellText (ConvertClipboardCellValue (value));
			}
			int upperBound = array.GetUpperBound (0);
			int upperBound2 = array.GetUpperBound (1);
			StringBuilder stringBuilder = new StringBuilder ();
			for (int i = 1; i <= upperBound; i++) {
				for (int j = 1; j <= upperBound2; j++) {
					if (j > 1) {
						stringBuilder.Append ('\t');
					}
					stringBuilder.Append (SanitizeClipboardCellText (ConvertClipboardCellValue (array [i, j])));
				}
				if (i < upperBound) {
					stringBuilder.Append ("\r\n");
				}
			}
			return stringBuilder.ToString ();
		}

		private static string ConvertClipboardCellValue (object value)
		{
			if (value == null) {
				return string.Empty;
			}
			if (value is bool flag) {
				return flag ? "TRUE" : "FALSE";
			}
			if (value is double num) {
				return num.ToString ("G15", CultureInfo.CurrentCulture);
			}
			return Convert.ToString (value, CultureInfo.CurrentCulture) ?? string.Empty;
		}

		private static string SanitizeClipboardCellText (string value)
		{
			return (value ?? string.Empty).Replace ("\r\n", " ").Replace ('\r', ' ').Replace ('\n', ' ')
				.Replace ('\t', ' ');
		}

		private static void WriteClipboardTextWithRetry (string clipboardText)
		{
			if (clipboardText == null) {
				clipboardText = string.Empty;
			}
			Exception ex = null;
			for (int i = 1; i <= 3; i++) {
				try {
					DataObject dataObject = new DataObject ();
					dataObject.SetData (DataFormats.UnicodeText, autoConvert: true, clipboardText);
					dataObject.SetData (DataFormats.Text, autoConvert: true, clipboardText);
					Clipboard.SetDataObject (dataObject, copy: true);
					return;
				} catch (ExternalException ex2) {
					ex = ex2;
					if (i >= 3) {
						throw;
					}
					Thread.Sleep (40);
				}
			}
			if (ex != null) {
				throw ex;
			}
		}

		private static string GetWorkbookKey (Workbook workbook)
		{
			if (workbook == null) {
				return string.Empty;
			}
			string text = workbook.FullName ?? string.Empty;
			return string.IsNullOrWhiteSpace (text) ? (workbook.Name ?? string.Empty) : text;
		}

		private static void ReleaseComObject (object comObject)
		{
			if (comObject == null) {
				return;
			}
			try {
				Marshal.FinalReleaseComObject (comObject);
			} catch {
			}
		}
	}
}
