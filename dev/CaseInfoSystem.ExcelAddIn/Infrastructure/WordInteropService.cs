using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class WordInteropService
	{
		internal sealed class WordInteropServiceTestHooks
		{
			internal Func<object, bool> IsWordApplicationAlive { get; set; }

			internal Func<string, object> GetActiveObject { get; set; }

			internal Func<string, Type> GetTypeFromProgID { get; set; }

			internal Func<Type, object> CreateInstance { get; set; }

			internal Func<string, string> ResolveToExistingLocalPath { get; set; }

			internal Func<string, string> NormalizePath { get; set; }

			internal Func<string, bool> FileExists { get; set; }
		}

		internal sealed class WordPerformanceState
		{
			internal bool HasScreenUpdating { get; set; }

			internal object ScreenUpdating { get; set; }

			internal bool HasDisplayAlerts { get; set; }

			internal object DisplayAlerts { get; set; }

			internal bool HasVisible { get; set; }

			internal object Visible { get; set; }
		}

		private const int WordFormatDocumentDefault = 16;

		private const int ShowWindowRestore = 9;

		private const int ShowWindowShow = 5;

		private static readonly IntPtr HwndTopMost = new IntPtr (-1);

		private static readonly IntPtr HwndNoTopMost = new IntPtr (-2);

		private const uint SwpNoMove = 2u;

		private const uint SwpNoSize = 1u;

		private const uint SwpShowWindow = 64u;

		private const int ForegroundRetryCount = 8;

		private const int ForegroundRetryDelayMs = 120;

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly Logger _logger;

		private readonly WordInteropServiceTestHooks _testHooks;

		private object _cachedWordApplication;

		[DllImport ("user32.dll")]
		private static extern bool SetForegroundWindow (IntPtr hWnd);

		[DllImport ("user32.dll")]
		private static extern bool ShowWindowAsync (IntPtr hWnd, int nCmdShow);

		[DllImport ("user32.dll")]
		private static extern bool IsIconic (IntPtr hWnd);

		[DllImport ("user32.dll")]
		private static extern bool BringWindowToTop (IntPtr hWnd);

		[DllImport ("user32.dll")]
		private static extern IntPtr GetForegroundWindow ();

		[DllImport ("user32.dll")]
		private static extern uint GetWindowThreadProcessId (IntPtr hWnd, out uint processId);

		[DllImport ("kernel32.dll")]
		private static extern uint GetCurrentThreadId ();

		[DllImport ("user32.dll")]
		private static extern bool AttachThreadInput (uint idAttach, uint idAttachTo, bool fAttach);

		[DllImport ("user32.dll")]
		private static extern bool SetWindowPos (IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

		internal WordInteropService (PathCompatibilityService pathCompatibilityService, Logger logger)
			: this (pathCompatibilityService, logger, testHooks: null)
		{
		}

		internal WordInteropService (PathCompatibilityService pathCompatibilityService, Logger logger, WordInteropServiceTestHooks testHooks)
		{
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
			_testHooks = testHooks;
		}

		internal object AcquireWordApplication (out bool createdNew)
		{
			createdNew = false;
			if (IsWordApplicationAliveSafe (_cachedWordApplication)) {
				return _cachedWordApplication;
			}
			_cachedWordApplication = null;
			object obj = null;
			try {
				obj = GetActiveObject ("Word.Application");
			} catch {
			}
			if (!IsWordApplicationAliveSafe (obj)) {
				obj = null;
			}
			if (obj == null) {
				try {
					Type typeFromProgID = GetTypeFromProgID ("Word.Application");
					if (typeFromProgID != null) {
						obj = CreateInstance (typeFromProgID);
						if (IsWordApplicationAliveSafe (obj)) {
							createdNew = true;
						} else {
							obj = null;
						}
					}
				} catch (Exception exception) {
					_logger.Error ("WordInteropService.AcquireWordApplication failed.", exception);
					obj = null;
				}
			}
			_cachedWordApplication = obj;
			return obj;
		}

		internal object EnsureWordApplication (ref object wordApplication)
		{
			if (IsWordApplicationAliveSafe (wordApplication)) {
				return wordApplication;
			}
			wordApplication = AcquireWordApplication (out var _);
			return wordApplication;
		}

		internal void WarmUpApplication ()
		{
			bool createdNew;
			object obj = AcquireWordApplication (out createdNew);
			if (obj == null) {
				return;
			}
			try {
				dynamic val = obj;
				if (Convert.ToInt32 (val.Documents.Count) == 0) {
					val.Visible = false;
				}
			} catch {
			}
			_logger.Info ("WordInteropService warm-up completed. createdNew=" + createdNew);
		}

		internal object CreateDocumentFromTemplate (object wordApplication, string templatePath)
		{
			if (wordApplication == null) {
				_logger.Warn ("WordInteropService.CreateDocumentFromTemplate skipped because wordApplication was null.");
				return null;
			}
			string text;
			try {
				text = NormalizePath (ResolveToExistingLocalPath (templatePath));
			} catch (Exception exception) {
				_logger.Error ("WordInteropService.CreateDocumentFromTemplate template path resolution failed.", exception);
				return null;
			}
			if (text.Length == 0 || !FileExistsSafe (text)) {
				_logger.Warn ("WordInteropService.CreateDocumentFromTemplate template not found. template=" + (templatePath ?? string.Empty) + " resolved=" + text);
				return null;
			}
			try {
				return ((dynamic)wordApplication).Documents.Add (text, false);
			} catch (Exception exception2) {
				_logger.Error ("WordInteropService.CreateDocumentFromTemplate Documents.Add failed.", exception2);
				return null;
			}
		}

		internal string SaveDocumentAsDocx (object wordDocument, string fullPath)
		{
			if (wordDocument == null) {
				throw new ArgumentNullException ("wordDocument");
			}
			string text = _pathCompatibilityService.NormalizePath (fullPath);
			if (text.Length == 0) {
				throw new InvalidOperationException ("保存先パスが解決されていません。");
			}
			string directoryName = Path.GetDirectoryName (text);
			if (!string.IsNullOrWhiteSpace (directoryName) && !_pathCompatibilityService.EnsureFolderSafe (directoryName)) {
				throw new InvalidOperationException ("保存先フォルダを作成できませんでした。");
			}
			((dynamic)wordDocument).SaveAs2 (text, 16);
			return text;
		}

		internal WordPerformanceState BeginPerformanceScope (object wordApplication, bool hideWhenNew, bool createdNewWord)
		{
			WordPerformanceState wordPerformanceState = new WordPerformanceState ();
			if (wordApplication == null) {
				return wordPerformanceState;
			}
			try {
				wordPerformanceState.HasScreenUpdating = true;
				wordPerformanceState.ScreenUpdating = (object)((dynamic)wordApplication).ScreenUpdating;
				wordPerformanceState.HasDisplayAlerts = true;
				wordPerformanceState.DisplayAlerts = (object)((dynamic)wordApplication).DisplayAlerts;
				wordPerformanceState.HasVisible = true;
				wordPerformanceState.Visible = (object)((dynamic)wordApplication).Visible;
				((dynamic)wordApplication).ScreenUpdating = false;
				((dynamic)wordApplication).DisplayAlerts = 0;
				if (hideWhenNew && createdNewWord) {
					((dynamic)wordApplication).Visible = false;
				}
			} catch {
			}
			return wordPerformanceState;
		}

		internal void RestorePerformanceState (object wordApplication, WordPerformanceState state)
		{
			if (wordApplication == null || state == null) {
				return;
			}
			try {
				if (state.HasScreenUpdating) {
					((dynamic)wordApplication).ScreenUpdating = state.ScreenUpdating;
				}
				if (state.HasDisplayAlerts) {
					((dynamic)wordApplication).DisplayAlerts = state.DisplayAlerts;
				}
				if (state.HasVisible) {
					((dynamic)wordApplication).Visible = state.Visible;
				}
			} catch {
			}
		}

		internal bool IsDocumentOpen (object wordApplication, string fullPath)
		{
			string text = _pathCompatibilityService.NormalizePath (fullPath).ToLowerInvariant ();
			if (wordApplication == null || text.Length == 0) {
				return false;
			}
			try {
				foreach (object item in ((dynamic)wordApplication).Documents) {
					dynamic val = item;
					string a = _pathCompatibilityService.NormalizePath (Convert.ToString (val.FullName)).ToLowerInvariant ();
					if (string.Equals (a, text, StringComparison.OrdinalIgnoreCase)) {
						return true;
					}
				}
			} catch {
			}
			return false;
		}

		internal void CloseDocumentNoSave (ref object wordDocument)
		{
			try {
				if (wordDocument != null) {
					dynamic val = wordDocument;
					val.Close (false);
				}
			} catch (Exception exception) {
				_logger.Error ("WordInteropService.CloseDocumentNoSave failed.", exception);
			} finally {
				ReleaseComObject (wordDocument);
				wordDocument = null;
			}
		}

		internal void ShowDocument (object wordApplication, object wordDocument)
		{
			if (wordApplication != null) {
				Stopwatch stopwatch = Stopwatch.StartNew ();
				Stopwatch stopwatch2 = Stopwatch.StartNew ();
				((dynamic)wordApplication).Visible = true;
				_logger.Debug ("WordInteropService.ShowDocument", "VisibleSet elapsed=" + FormatElapsedSeconds (stopwatch2.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (stopwatch.Elapsed));
				stopwatch2.Restart ();
				TryRestoreWordWindow ((dynamic)wordApplication);
				_logger.Debug ("WordInteropService.ShowDocument", "WindowRestoreRequested elapsed=" + FormatElapsedSeconds (stopwatch2.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (stopwatch.Elapsed));
				stopwatch2.Restart ();
				if (wordDocument != null) {
					((dynamic)wordDocument).Activate ();
				}
				_logger.Debug ("WordInteropService.ShowDocument", "DocumentActivated elapsed=" + FormatElapsedSeconds (stopwatch2.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (stopwatch.Elapsed) + " hasDocument=" + (wordDocument != null));
				stopwatch2.Restart ();
				((dynamic)wordApplication).Activate ();
				_logger.Debug ("WordInteropService.ShowDocument", "ApplicationActivated elapsed=" + FormatElapsedSeconds (stopwatch2.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (stopwatch.Elapsed));
				stopwatch2.Restart ();
				TryBringWordToFront ((dynamic)wordApplication);
				_logger.Debug ("WordInteropService.ShowDocument", "BringToFrontRequested elapsed=" + FormatElapsedSeconds (stopwatch2.Elapsed) + " totalElapsed=" + FormatElapsedSeconds (stopwatch.Elapsed));
			}
		}

		internal void QuitApplicationNoSave (ref object wordApplication)
		{
			try {
				if (wordApplication != null) {
					dynamic val = wordApplication;
					val.Quit (false);
				}
			} catch (Exception exception) {
				_logger.Error ("WordInteropService.QuitApplicationNoSave failed.", exception);
			} finally {
				if (_cachedWordApplication == wordApplication) {
					_cachedWordApplication = null;
				}
				ReleaseComObject (wordApplication);
				wordApplication = null;
			}
		}

		private static bool IsWordApplicationAlive (object wordApplication)
		{
			if (wordApplication == null) {
				return false;
			}
			try {
				object obj = ((dynamic)wordApplication).Name;
				return obj != null;
			} catch {
				return false;
			}
		}

		private bool IsWordApplicationAliveSafe (object wordApplication)
		{
			if (_testHooks != null && _testHooks.IsWordApplicationAlive != null) {
				return _testHooks.IsWordApplicationAlive (wordApplication);
			}
			return IsWordApplicationAlive (wordApplication);
		}

		private object GetActiveObject (string progId)
		{
			if (_testHooks != null && _testHooks.GetActiveObject != null) {
				return _testHooks.GetActiveObject (progId);
			}
			return Marshal.GetActiveObject (progId);
		}

		private Type GetTypeFromProgID (string progId)
		{
			return (_testHooks != null && _testHooks.GetTypeFromProgID != null) ? _testHooks.GetTypeFromProgID (progId) : Type.GetTypeFromProgID (progId);
		}

		private object CreateInstance (Type type)
		{
			return (_testHooks != null && _testHooks.CreateInstance != null) ? _testHooks.CreateInstance (type) : Activator.CreateInstance (type);
		}

		private string ResolveToExistingLocalPath (string path)
		{
			return (_testHooks != null && _testHooks.ResolveToExistingLocalPath != null) ? _testHooks.ResolveToExistingLocalPath (path) : _pathCompatibilityService.ResolveToExistingLocalPath (path);
		}

		private string NormalizePath (string path)
		{
			return (_testHooks != null && _testHooks.NormalizePath != null) ? _testHooks.NormalizePath (path) : _pathCompatibilityService.NormalizePath (path);
		}

		private bool FileExistsSafe (string path)
		{
			return (_testHooks != null && _testHooks.FileExists != null) ? _testHooks.FileExists (path) : _pathCompatibilityService.FileExistsSafe (path);
		}

		private static string FormatElapsedSeconds (TimeSpan elapsed)
		{
			return elapsed.TotalSeconds.ToString ("0.000");
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

		private static void TryRestoreWordWindow (dynamic wordApplication)
		{
			try {
				IntPtr intPtr = WordInteropService.GetPrimaryWordWindowHandle (wordApplication);
				if (!(intPtr == IntPtr.Zero)) {
					ShowWindowAsync (intPtr, IsIconic (intPtr) ? 9 : 5);
				}
			} catch {
			}
		}

		private static void TryBringWordToFront (dynamic wordApplication)
		{
			try {
				IntPtr intPtr = WordInteropService.GetPrimaryWordWindowHandle (wordApplication);
				if (!(intPtr == IntPtr.Zero)) {
					for (int i = 0; i < 8; i++) {
						TryBringWindowToFront (intPtr);
						Thread.Sleep (120);
					}
				}
			} catch {
			}
		}

		private static IntPtr GetPrimaryWordWindowHandle (dynamic wordApplication)
		{
			IntPtr intPtr = WordInteropService.GetActiveWordWindowHandle (wordApplication);
			if (intPtr != IntPtr.Zero) {
				return intPtr;
			}
			try {
				return new IntPtr ((int)wordApplication.Hwnd);
			} catch {
				return IntPtr.Zero;
			}
		}

		private static IntPtr GetActiveWordWindowHandle (dynamic wordApplication)
		{
			try {
				dynamic val = wordApplication.ActiveWindow;
				if (val == null) {
					return IntPtr.Zero;
				}
				return new IntPtr ((int)val.Hwnd);
			} catch {
				return IntPtr.Zero;
			}
		}

		private static void TryBringWindowToFront (IntPtr hwnd)
		{
			if (hwnd == IntPtr.Zero) {
				return;
			}
			IntPtr foregroundWindow = GetForegroundWindow ();
			uint processId;
			uint num = ((!(foregroundWindow == IntPtr.Zero)) ? GetWindowThreadProcessId (foregroundWindow, out processId) : 0u);
			uint currentThreadId = GetCurrentThreadId ();
			bool flag = false;
			try {
				if (num != 0 && num != currentThreadId) {
					flag = AttachThreadInput (currentThreadId, num, fAttach: true);
				}
				ShowWindowAsync (hwnd, IsIconic (hwnd) ? 9 : 5);
				SetWindowPos (hwnd, HwndTopMost, 0, 0, 0, 0, 67u);
				BringWindowToTop (hwnd);
				SetForegroundWindow (hwnd);
				SetWindowPos (hwnd, HwndNoTopMost, 0, 0, 0, 0, 67u);
			} finally {
				if (flag) {
					AttachThreadInput (currentThreadId, num, fAttach: false);
				}
			}
		}
	}
}



