using System;
using System.Collections.Generic;
using System.Runtime.ExceptionServices;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class ExcelApplicationStateScope : IDisposable
	{
		private enum StateKind
		{
			ScreenUpdating,
			EnableEvents,
			DisplayAlerts,
			Calculation
		}

		private readonly Application _application;

		private readonly bool _suppressRestoreExceptions;

		private readonly bool _screenUpdating;

		private readonly bool _enableEvents;

		private readonly bool _displayAlerts;

		private readonly XlCalculation _calculation;

		private readonly List<StateKind> _restoreOrder = new List<StateKind> ();

		private bool _screenUpdatingChanged;

		private bool _enableEventsChanged;

		private bool _displayAlertsChanged;

		private bool _calculationChanged;

		private bool _disposed;

		internal ExcelApplicationStateScope (Application application, bool suppressRestoreExceptions = false)
		{
			_application = application;
			_suppressRestoreExceptions = suppressRestoreExceptions;
			if (_application == null) {
				return;
			}
			_screenUpdating = _application.ScreenUpdating;
			_enableEvents = _application.EnableEvents;
			_displayAlerts = _application.DisplayAlerts;
			_calculation = _application.Calculation;
		}

		internal void SetScreenUpdating (bool value)
		{
			if (_application == null) {
				return;
			}
			MarkChanged (StateKind.ScreenUpdating, ref _screenUpdatingChanged);
			_application.ScreenUpdating = value;
		}

		internal void SetEnableEvents (bool value)
		{
			if (_application == null) {
				return;
			}
			MarkChanged (StateKind.EnableEvents, ref _enableEventsChanged);
			_application.EnableEvents = value;
		}

		internal void SetDisplayAlerts (bool value)
		{
			if (_application == null) {
				return;
			}
			MarkChanged (StateKind.DisplayAlerts, ref _displayAlertsChanged);
			_application.DisplayAlerts = value;
		}

		internal void SetCalculation (XlCalculation value)
		{
			if (_application == null) {
				return;
			}
			MarkChanged (StateKind.Calculation, ref _calculationChanged);
			_application.Calculation = value;
		}

		public void Dispose ()
		{
			if (_disposed) {
				return;
			}
			_disposed = true;
			if (_application == null || _restoreOrder.Count == 0) {
				return;
			}
			ExceptionDispatchInfo exceptionDispatchInfo = null;
			for (int i = 0; i < _restoreOrder.Count; i++) {
				try {
					RestoreState (_restoreOrder [i]);
				} catch (Exception ex) {
					if (exceptionDispatchInfo == null) {
						exceptionDispatchInfo = ExceptionDispatchInfo.Capture (ex);
					}
				}
			}
			if (!_suppressRestoreExceptions && exceptionDispatchInfo != null) {
				exceptionDispatchInfo.Throw ();
			}
		}

		private void MarkChanged (StateKind stateKind, ref bool changed)
		{
			if (changed) {
				return;
			}
			changed = true;
			_restoreOrder.Add (stateKind);
		}

		private void RestoreState (StateKind stateKind)
		{
			switch (stateKind) {
			case StateKind.ScreenUpdating:
				_application.ScreenUpdating = _screenUpdating;
				break;
			case StateKind.EnableEvents:
				_application.EnableEvents = _enableEvents;
				break;
			case StateKind.DisplayAlerts:
				_application.DisplayAlerts = _displayAlerts;
				break;
			case StateKind.Calculation:
				_application.Calculation = _calculation;
				break;
			}
		}
	}
}
