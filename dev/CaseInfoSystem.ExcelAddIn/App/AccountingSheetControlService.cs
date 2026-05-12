using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class AccountingSheetControlService
	{
		private struct CheckboxState
		{
			internal string Y15State { get; }

			internal string Y16State { get; }

			internal CheckboxState (bool y15Checked, bool y16Checked)
			{
				Y15State = (y15Checked ? "1" : "0");
				Y16State = (y16Checked ? "1" : "0");
			}

			internal bool Equals (CheckboxState other)
			{
				return string.Equals (Y15State, other.Y15State, StringComparison.Ordinal) && string.Equals (Y16State, other.Y16State, StringComparison.Ordinal);
			}

			internal string ToLogText ()
			{
				return "Y15=" + Y15State + ",Y16=" + Y16State;
			}
		}

		private const string CheckedStateTrue = "1";

		private const string CheckedStateFalse = "0";

		private const string LinkedCellCheck2 = "Y15";

		private const string LinkedCellCheck3 = "Y16";

		private const string LinkedCellWithholding = "Y24";

		private const string VisualCellCheck2 = "Z15";

		private const string VisualCellCheck3 = "Z16";

		private const string VisualCellWithholding = "Z24";

		private const string CheckedMark = "☑";

		private const string UncheckedMark = "□";

		private const string LawyerSourceRangeAddress = "A41:A44";

		private const string PaymentHistoryLawyerRangeAddress = "A6:A9";

		private readonly WorkbookRoleResolver _workbookRoleResolver;

		private readonly AccountingWorkbookService _accountingWorkbookService;

		private readonly Logger _logger;

		private readonly Dictionary<string, CheckboxState> _checkboxStates;

		private readonly Dictionary<string, string> _lastNonCheckboxSelectionAddresses;

		private readonly HashSet<string> _configuredWorkbookKeys;

		private readonly HashSet<string> _configuringWorkbookKeys;

		private readonly HashSet<string> _lawyerReflectionWorkbookKeys;

		private bool _suppressSelectionChangeHandling;

		private bool _suppressAfterCalculateHandling;

		private bool _suppressCheckboxEventHandling;

		internal AccountingSheetControlService (WorkbookRoleResolver workbookRoleResolver, AccountingWorkbookService accountingWorkbookService, Logger logger)
		{
			_workbookRoleResolver = workbookRoleResolver ?? throw new ArgumentNullException ("workbookRoleResolver");
			_accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException ("accountingWorkbookService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
			_checkboxStates = new Dictionary<string, CheckboxState> (StringComparer.OrdinalIgnoreCase);
			_lastNonCheckboxSelectionAddresses = new Dictionary<string, string> (StringComparer.OrdinalIgnoreCase);
			_configuredWorkbookKeys = new HashSet<string> (StringComparer.OrdinalIgnoreCase);
			_configuringWorkbookKeys = new HashSet<string> (StringComparer.OrdinalIgnoreCase);
			_lawyerReflectionWorkbookKeys = new HashSet<string> (StringComparer.OrdinalIgnoreCase);
		}

		internal void HandleSheetChange (object sheetObject, Range target)
		{
			if (!(sheetObject is Worksheet worksheet) || target == null) {
				_logger.Debug ("AccountingSheetControlService.HandleSheetChange", "skipped. worksheetOrTargetMissing=true");
				return;
			}
			Workbook workbook = null;
			try {
				workbook = worksheet.Parent as Workbook;
				string text = ((workbook == null) ? string.Empty : (workbook.Name ?? string.Empty));
				string text2 = worksheet.CodeName ?? string.Empty;
				string text3 = SafeAddress (target);
				_logger.Debug ("AccountingSheetControlService.HandleSheetChange", "received. workbook=" + text + ", sheet=" + text2 + ", target=" + text3);
				if (!_workbookRoleResolver.IsAccountingWorkbook (workbook)) {
					_logger.Debug ("AccountingSheetControlService.HandleSheetChange", "ignored. reason=notAccountingWorkbook, workbook=" + text);
				} else if (_suppressCheckboxEventHandling) {
					_logger.Debug ("AccountingSheetControlService.HandleSheetChange", "ignored. reason=suppressedProgrammaticCheckboxUpdate, sheet=" + text2 + ", target=" + text3);
				} else if (!IsMainAccountingFormSheet (text2)) {
					_logger.Debug ("AccountingSheetControlService.HandleSheetChange", "ignored. reason=notMainFormSheet, sheet=" + text2);
				} else if (string.Equals (text3, LinkedCellCheck2, StringComparison.OrdinalIgnoreCase)) {
					_logger.Debug ("AccountingSheetControlService.HandleSheetChange", "matched. sheet=" + text2 + ", triggerCell=" + LinkedCellCheck2);
					UpdateTrackedCheckboxState (workbook, text2);
					SyncCheckboxVisuals (workbook, text2);
					ApplyBaseAmountHighlight (workbook, text2, LinkedCellCheck2, LinkedCellCheck3);
				} else if (string.Equals (text3, LinkedCellCheck3, StringComparison.OrdinalIgnoreCase)) {
					_logger.Debug ("AccountingSheetControlService.HandleSheetChange", "matched. sheet=" + text2 + ", triggerCell=" + LinkedCellCheck3);
					UpdateTrackedCheckboxState (workbook, text2);
					SyncCheckboxVisuals (workbook, text2);
					ApplyBaseAmountHighlight (workbook, text2, LinkedCellCheck3, LinkedCellCheck2);
				} else if (string.Equals (text3, LinkedCellWithholding, StringComparison.OrdinalIgnoreCase)) {
					_logger.Debug ("AccountingSheetControlService.HandleSheetChange", "matched. sheet=" + text2 + ", triggerCell=" + LinkedCellWithholding);
					UpdateTrackedCheckboxState (workbook, text2);
					SyncCheckboxVisuals (workbook, text2);
				} else if (IntersectsRange (worksheet, target, "A41:A44")) {
					string workbookKey = GetWorkbookKey (workbook);
					if (_lawyerReflectionWorkbookKeys.Contains (workbookKey)) {
						_logger.Debug ("AccountingSheetControlService.HandleSheetChange", "ignored. reason=lawyerReflectionReentry, workbook=" + workbookKey + ", sheet=" + text2);
					} else {
						ReflectLawyersAcrossAccountingSheets (workbook, text2);
					}
				} else {
					_logger.Debug ("AccountingSheetControlService.HandleSheetChange", "ignored. reason=unmatchedTarget, sheet=" + text2 + ", target=" + text3);
				}
			} catch (Exception exception) {
				_logger.Error ("AccountingSheetControlService.HandleSheetChange failed.", exception);
			}
		}

		internal void HandleAfterCalculate (Microsoft.Office.Interop.Excel.Application application)
		{
			if (application == null) {
				_logger.Debug ("AccountingSheetControlService", "HandleAfterCalculate enter. guard=ApplicationNull");
				return;
			}
			if (_suppressAfterCalculateHandling || _suppressCheckboxEventHandling) {
				_logger.Debug ("AccountingSheetControlService", "HandleAfterCalculate enter. guard=Suppressed, suppressAfterCalculate=" + _suppressAfterCalculateHandling + ", suppressCheckbox=" + _suppressCheckboxEventHandling);
				return;
			}
			_logger.Debug ("AccountingSheetControlService", "HandleAfterCalculate enter. guard=Passed");
			if (IsCutCopyInProgress (application)) {
				_logger.Debug ("AccountingSheetControlService", "HandleAfterCalculate branch=SkippedCutCopyMode");
				return;
			}
			Workbook workbook = null;
			try {
				_suppressAfterCalculateHandling = true;
				workbook = application.ActiveWorkbook;
				Worksheet activeWorksheet = null;
				string workbookKey = GetWorkbookKey (workbook);
				string activeSheetName = string.Empty;
				string activeWindowCaption = string.Empty;
				try {
					activeWorksheet = application.ActiveSheet as Worksheet;
					activeSheetName = activeWorksheet?.CodeName ?? activeWorksheet?.Name ?? string.Empty;
				} catch {
					activeSheetName = string.Empty;
				}
				try {
					activeWindowCaption = application.ActiveWindow?.Caption as string ?? string.Empty;
				} catch {
					activeWindowCaption = string.Empty;
				}
				bool isAccountingWorkbook = _workbookRoleResolver.IsAccountingWorkbook (workbook);
				bool shouldSuspend = isAccountingWorkbook && ShouldSuspendForActiveSheet (application, workbook, activeWorksheet);
				_logger.Info ("AccountingSheetControlService HandleAfterCalculate target workbook=" + workbookKey + ", window=" + activeWindowCaption + ", activeSheet=" + activeSheetName + ", role=" + (isAccountingWorkbook ? "Accounting" : "Other") + ", shouldSuspend=" + shouldSuspend);
				if (isAccountingWorkbook && !shouldSuspend) {
					EnsureVstoManagedControls (workbook);
					TrackWorkbookCheckboxChanges (workbook);
					_logger.Info ("AccountingSheetControlService HandleAfterCalculate completed. workbook=" + workbookKey + ", branch=AccountingHandled");
				} else {
					_logger.Info ("AccountingSheetControlService HandleAfterCalculate completed. workbook=" + workbookKey + ", branch=" + (isAccountingWorkbook ? "Suspended" : "NonAccountingWorkbook"));
				}
			} catch (Exception exception) {
				_logger.Error ("AccountingSheetControlService.HandleAfterCalculate failed.", exception);
			} finally {
				_suppressAfterCalculateHandling = false;
				ReleaseWorkbookIterationObject (workbook);
			}
		}

		internal void HandleSheetActivated (object sheetObject)
		{
			if (!(sheetObject is Worksheet worksheet)) {
				return;
			}
			Workbook workbook = null;
			try {
				workbook = worksheet.Parent as Workbook;
				if (!_workbookRoleResolver.IsAccountingWorkbook (workbook)) {
					return;
				}
				if (IsCutCopyInProgress (workbook?.Application)) {
					_logger.Debug ("AccountingSheetControlService", "HandleSheetActivated skipped because cut/copy mode is active. workbook=" + GetWorkbookKey (workbook));
					return;
				}
				string sheetName = worksheet.CodeName ?? string.Empty;
				if (IsMainAccountingFormSheet (sheetName)) {
					SyncCheckboxVisuals (workbook, sheetName);
				}
			} catch (Exception exception) {
				_logger.Error ("AccountingSheetControlService.HandleSheetActivated failed.", exception);
			}
		}

		internal void HandleSheetSelectionChange (object sheetObject, Range target)
		{
			if (!(sheetObject is Worksheet worksheet) || target == null) {
				return;
			}
			Workbook workbook = null;
			try {
				workbook = worksheet.Parent as Workbook;
				if (!_workbookRoleResolver.IsAccountingWorkbook (workbook) || _suppressSelectionChangeHandling) {
					return;
				}
				string sheetName = worksheet.CodeName ?? string.Empty;
				bool flag = !IsMainAccountingFormSheet (sheetName);
				if (!flag && !((flag | ((dynamic)target.CountLarge != 1)) ? true : false)) {
					string text = SafeAddress (target);
					if (!IsVisualCheckboxCell (text)) {
						RememberNonCheckboxSelection (workbook, sheetName, text);
					} else if (string.Equals (text, VisualCellCheck2, StringComparison.OrdinalIgnoreCase)) {
						ToggleLinkedCheckbox (workbook, sheetName, LinkedCellCheck2, VisualCellCheck2);
						RestoreSelectionAfterCheckboxToggle (workbook, sheetName);
					} else if (string.Equals (text, VisualCellCheck3, StringComparison.OrdinalIgnoreCase)) {
						ToggleLinkedCheckbox (workbook, sheetName, LinkedCellCheck3, VisualCellCheck3);
						RestoreSelectionAfterCheckboxToggle (workbook, sheetName);
					} else if (string.Equals (text, VisualCellWithholding, StringComparison.OrdinalIgnoreCase)) {
						ToggleLinkedCheckbox (workbook, sheetName, LinkedCellWithholding, VisualCellWithholding);
						RestoreSelectionAfterCheckboxToggle (workbook, sheetName);
					}
				}
			} catch (Exception exception) {
				_logger.Error ("AccountingSheetControlService.HandleSheetSelectionChange failed.", exception);
			}
		}

		internal void RemoveWorkbookState (Workbook workbook)
		{
			if (workbook == null) {
				return;
			}
			string workbookKey = GetWorkbookKey (workbook);
			List<string> list = new List<string> ();
			foreach (KeyValuePair<string, CheckboxState> checkboxState in _checkboxStates) {
				if (checkboxState.Key.StartsWith (workbookKey + "|", StringComparison.OrdinalIgnoreCase)) {
					list.Add (checkboxState.Key);
				}
			}
			foreach (string item in list) {
				_checkboxStates.Remove (item);
			}
			List<string> list2 = new List<string> ();
			foreach (KeyValuePair<string, string> lastNonCheckboxSelectionAddress in _lastNonCheckboxSelectionAddresses) {
				if (lastNonCheckboxSelectionAddress.Key.StartsWith (workbookKey + "|", StringComparison.OrdinalIgnoreCase)) {
					list2.Add (lastNonCheckboxSelectionAddress.Key);
				}
			}
			foreach (string item2 in list2) {
				_lastNonCheckboxSelectionAddresses.Remove (item2);
			}
			_configuredWorkbookKeys.Remove (workbookKey);
			_configuringWorkbookKeys.Remove (workbookKey);
			_logger.Debug ("AccountingSheetControlService", "RemoveWorkbookState workbook=" + workbookKey + ", removedKeys=" + list.Count);
		}

		internal void EnsureVstoManagedControls (Workbook workbook)
		{
			if (!_workbookRoleResolver.IsAccountingWorkbook (workbook)) {
				return;
			}
			string workbookKey = GetWorkbookKey (workbook);
			if (_configuredWorkbookKeys.Contains (workbookKey) || _configuringWorkbookKeys.Contains (workbookKey)) {
				return;
			}
			_configuringWorkbookKeys.Add (workbookKey);
			try {
				foreach (string mainFormSheet in GetMainFormSheets ()) {
					SyncCheckboxVisuals (workbook, mainFormSheet);
				}
				_configuredWorkbookKeys.Add (workbookKey);
				_logger.Info ("Accounting cell checkbox controls synchronized for VSTO management. workbook=" + workbookKey);
			} finally {
				_configuringWorkbookKeys.Remove (workbookKey);
			}
		}

		private bool ShouldSuspendForActiveSheet (Microsoft.Office.Interop.Excel.Application application, Workbook workbook, Worksheet activeWorksheet)
		{
			if (application == null || workbook == null || !_workbookRoleResolver.IsAccountingWorkbook (workbook)) {
				return false;
			}
			try {
				Worksheet worksheet = activeWorksheet;
				if (worksheet == null) {
					worksheet = application.ActiveSheet as Worksheet;
				}
				if (worksheet == null) {
					return false;
				}
				Workbook workbook2 = worksheet.Parent as Workbook;
				if (workbook2 != workbook) {
					return false;
				}
				string text = worksheet.CodeName ?? worksheet.Name ?? string.Empty;
				bool flag = string.Equals (text, "お支払い履歴", StringComparison.OrdinalIgnoreCase);
				if (flag) {
					_logger.Debug ("AccountingSheetControlService", "HandleAfterCalculate suspended. workbook=" + GetWorkbookKey (workbook) + ", activeSheet=" + text);
				}
				return flag;
			} catch (Exception exception) {
				_logger.Error ("AccountingSheetControlService.ShouldSuspendForActiveSheet failed.", exception);
				return false;
			}
		}

		private void TrackWorkbookCheckboxChanges (Workbook workbook)
		{
			if (!_workbookRoleResolver.IsAccountingWorkbook (workbook)) {
				return;
			}
			foreach (string mainFormSheet in GetMainFormSheets ()) {
				TrackSheetCheckboxChanges (workbook, mainFormSheet);
			}
		}

		private void TrackSheetCheckboxChanges (Workbook workbook, string sheetName)
		{
			string key = BuildStateKey (workbook, sheetName);
			CheckboxState checkboxState = new CheckboxState (ReadLinkedBoolean (workbook, sheetName, LinkedCellCheck2), ReadLinkedBoolean (workbook, sheetName, LinkedCellCheck3));
			if (!_checkboxStates.TryGetValue (key, out var value)) {
				_checkboxStates [key] = checkboxState;
				_logger.Debug ("AccountingSheetControlService", "TrackSheetCheckboxChanges initialized. sheet=" + sheetName + ", y15=" + checkboxState.Y15State + ", y16=" + checkboxState.Y16State);
			} else if (!value.Equals (checkboxState)) {
				_logger.Debug ("AccountingSheetControlService", "TrackSheetCheckboxChanges changed. sheet=" + sheetName + ", previous=" + value.ToLogText () + ", current=" + checkboxState.ToLogText ());
				_checkboxStates [key] = checkboxState;
				if (!string.Equals (value.Y15State, checkboxState.Y15State, StringComparison.Ordinal)) {
					SyncCheckboxVisuals (workbook, sheetName);
					ApplyBaseAmountHighlight (workbook, sheetName, LinkedCellCheck2, LinkedCellCheck3);
				}
				if (!string.Equals (value.Y16State, checkboxState.Y16State, StringComparison.Ordinal)) {
					SyncCheckboxVisuals (workbook, sheetName);
					ApplyBaseAmountHighlight (workbook, sheetName, LinkedCellCheck3, LinkedCellCheck2);
				}
			}
		}

		private void ApplyBaseAmountHighlight (Workbook workbook, string sheetName, string currentCheckboxCell, string otherCheckboxCell)
		{
			bool flag = ReadLinkedBoolean (workbook, sheetName, currentCheckboxCell);
			bool flag2 = ReadLinkedBoolean (workbook, sheetName, otherCheckboxCell);
			string value = _accountingWorkbookService.ReadText (workbook, sheetName, "F33");
			if (flag && string.IsNullOrWhiteSpace (value)) {
				_accountingWorkbookService.SetInteriorColorIndex (workbook, sheetName, "F33", 36);
				UserErrorService.ShowOkNotification ("シート下部の「付記事項欄」の黄色エリアに\r\n経済的利益額を入力してください。", "案件情報System", MessageBoxIcon.Asterisk);
				_logger.Info ("Accounting checkbox highlight applied. sheet=" + sheetName + ", triggerCell=" + currentCheckboxCell + ", messageSource=VSTO");
			} else if (!flag2 || !string.IsNullOrWhiteSpace (value)) {
				_accountingWorkbookService.SetInteriorColorIndex (workbook, sheetName, "F33", 0);
				_logger.Info ("Accounting checkbox highlight cleared. sheet=" + sheetName + ", triggerCell=" + currentCheckboxCell);
			}
		}

		private static bool IsMainAccountingFormSheet (string sheetName)
		{
			return string.Equals (sheetName, "見積書", StringComparison.OrdinalIgnoreCase) || string.Equals (sheetName, "請求書", StringComparison.OrdinalIgnoreCase) || string.Equals (sheetName, "領収書", StringComparison.OrdinalIgnoreCase) || string.Equals (sheetName, "会計依頼書", StringComparison.OrdinalIgnoreCase);
		}

		private bool ReadLinkedBoolean (Workbook workbook, string sheetName, string address)
		{
			string a = _accountingWorkbookService.ReadText (workbook, sheetName, address);
			return string.Equals (a, "TRUE", StringComparison.OrdinalIgnoreCase) || string.Equals (a, "1", StringComparison.OrdinalIgnoreCase) || string.Equals (a, "ON", StringComparison.OrdinalIgnoreCase) || string.Equals (a, "-4146", StringComparison.OrdinalIgnoreCase);
		}

		private void ToggleLinkedCheckbox (Workbook workbook, string sheetName, string linkedCellAddress, string visualCellAddress)
		{
			bool flag = ReadLinkedBoolean (workbook, sheetName, linkedCellAddress);
			bool flag2 = !flag;
			try {
				_suppressCheckboxEventHandling = true;
				_accountingWorkbookService.WriteCellValue (workbook, sheetName, linkedCellAddress, flag2);
				_accountingWorkbookService.WriteCellValue (workbook, sheetName, visualCellAddress, flag2 ? "☑" : "□");
				UpdateTrackedCheckboxState (workbook, sheetName);
			} finally {
				_suppressCheckboxEventHandling = false;
			}
			if (string.Equals (linkedCellAddress, LinkedCellCheck2, StringComparison.OrdinalIgnoreCase)) {
				ApplyBaseAmountHighlight (workbook, sheetName, LinkedCellCheck2, LinkedCellCheck3);
			} else if (string.Equals (linkedCellAddress, LinkedCellCheck3, StringComparison.OrdinalIgnoreCase)) {
				ApplyBaseAmountHighlight (workbook, sheetName, LinkedCellCheck3, LinkedCellCheck2);
			}
			_logger.Info ("Accounting cell checkbox toggled. sheet=" + sheetName + ", visualCell=" + visualCellAddress + ", linkedCell=" + linkedCellAddress + ", nextState=" + flag2);
		}

		private void SyncCheckboxVisuals (Workbook workbook, string sheetName)
		{
			try {
				_suppressCheckboxEventHandling = true;
				_accountingWorkbookService.WriteCellValue (workbook, sheetName, VisualCellCheck2, ReadLinkedBoolean (workbook, sheetName, LinkedCellCheck2) ? CheckedMark : UncheckedMark);
				_accountingWorkbookService.WriteCellValue (workbook, sheetName, VisualCellCheck3, ReadLinkedBoolean (workbook, sheetName, LinkedCellCheck3) ? CheckedMark : UncheckedMark);
				_accountingWorkbookService.WriteCellValue (workbook, sheetName, VisualCellWithholding, ReadLinkedBoolean (workbook, sheetName, LinkedCellWithholding) ? CheckedMark : UncheckedMark);
			} finally {
				_suppressCheckboxEventHandling = false;
			}
		}

		private void UpdateTrackedCheckboxState (Workbook workbook, string sheetName)
		{
			_checkboxStates [BuildStateKey (workbook, sheetName)] = new CheckboxState (ReadLinkedBoolean (workbook, sheetName, LinkedCellCheck2), ReadLinkedBoolean (workbook, sheetName, LinkedCellCheck3));
		}

		private void RememberNonCheckboxSelection (Workbook workbook, string sheetName, string address)
		{
			if (!string.IsNullOrWhiteSpace (address) && !IsVisualCheckboxCell (address)) {
				_lastNonCheckboxSelectionAddresses [BuildStateKey (workbook, sheetName)] = address;
			}
		}

		private void RestoreSelectionAfterCheckboxToggle (Workbook workbook, string sheetName)
		{
			if (!_lastNonCheckboxSelectionAddresses.TryGetValue (BuildStateKey (workbook, sheetName), out var value) || string.IsNullOrWhiteSpace (value) || IsVisualCheckboxCell (value)) {
				value = "A1";
			}
			try {
				_suppressSelectionChangeHandling = true;
				_accountingWorkbookService.ActivateCell (workbook, sheetName, value);
			} finally {
				_suppressSelectionChangeHandling = false;
			}
		}

		private static bool IsVisualCheckboxCell (string address)
		{
			return string.Equals (address, VisualCellCheck2, StringComparison.OrdinalIgnoreCase) || string.Equals (address, VisualCellCheck3, StringComparison.OrdinalIgnoreCase) || string.Equals (address, VisualCellWithholding, StringComparison.OrdinalIgnoreCase);
		}

		private static bool IsCutCopyInProgress (Microsoft.Office.Interop.Excel.Application application)
		{
			if (application == null) {
				return false;
			}
			XlCutCopyMode cutCopyMode = application.CutCopyMode;
			return cutCopyMode == XlCutCopyMode.xlCopy || cutCopyMode == XlCutCopyMode.xlCut;
		}

		private void ReflectLawyersAcrossAccountingSheets (Workbook workbook, string sourceSheetName)
		{
			string workbookKey = GetWorkbookKey (workbook);
			_lawyerReflectionWorkbookKeys.Add (workbookKey);
			try {
				foreach (string mainFormSheet in GetMainFormSheets ()) {
					if (!string.Equals (mainFormSheet, sourceSheetName, StringComparison.OrdinalIgnoreCase)) {
						_accountingWorkbookService.CopyValueRange (workbook, sourceSheetName, "A41:A44", mainFormSheet, "A41:A44");
					}
				}
				_accountingWorkbookService.CopyValueRange (workbook, sourceSheetName, "A41:A44", "お支払い履歴", "A6:A9");
				_logger.Info ("Accounting lawyers reflected from sheet change. source=" + sourceSheetName);
			} finally {
				_lawyerReflectionWorkbookKeys.Remove (workbookKey);
			}
		}

		private static bool IntersectsRange (Worksheet worksheet, Range target, string address)
		{
			Range range = null;
			Range range2 = null;
			try {
				range = ((_Worksheet)worksheet)?.get_Range ((object)address, Type.Missing);
				if (range == null || target == null) {
					return false;
				}
				range2 = ((worksheet.Application == null) ? null : worksheet.Application.Intersect (target, range, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing));
				return range2 != null;
			} finally {
				ReleaseWorkbookIterationObject (range2);
				ReleaseWorkbookIterationObject (range);
			}
		}

		private static IEnumerable<string> GetMainFormSheets ()
		{
			yield return "見積書";
			yield return "請求書";
			yield return "領収書";
			yield return "会計依頼書";
		}

		private static string BuildStateKey (Workbook workbook, string sheetName)
		{
			return GetWorkbookKey (workbook) + "|" + (sheetName ?? string.Empty);
		}

		private static string GetWorkbookKey (Workbook workbook)
		{
			if (workbook == null) {
				return string.Empty;
			}
			string text = workbook.FullName ?? string.Empty;
			return string.IsNullOrWhiteSpace (text) ? (workbook.Name ?? string.Empty) : text;
		}

		private static void ReleaseWorkbookIterationObject (object comObject)
		{
			CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.Release (comObject);
		}

		private static string SafeAddress (Range range)
		{
			try {
				return (range == null) ? string.Empty : (Convert.ToString (range.get_Address ((object)false, (object)false, XlReferenceStyle.xlA1, Type.Missing, Type.Missing)) ?? string.Empty);
			} catch {
				return string.Empty;
			}
		}
	}
}
