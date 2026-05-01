using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class FolderWindowService
	{
		private const int WaitForFolderWindowTimeoutMs = 1500;

		private const int WaitForFolderWindowIntervalMs = 100;

		internal sealed class TestHooks
		{
			internal Func<string, string, bool> StartFolderProcess { get; set; }

			internal Func<string, IntPtr> TryFindExplorerWindow { get; set; }
		}

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly Logger _logger;

		internal TestHooks Hooks { get; set; }

		internal FolderWindowService (PathCompatibilityService pathCompatibilityService, Logger logger)
		{
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal IntPtr OpenFolderAndWait (string folderPath, string reason)
		{
			string text = _pathCompatibilityService.NormalizePath (folderPath);
			if (string.IsNullOrWhiteSpace (text)) {
				return IntPtr.Zero;
			}
			try {
				if (!OpenFolderCore (text, reason)) {
					return IntPtr.Zero;
				}
				IntPtr intPtr = WaitForFolderWindow (text);
				_logger.Info ("Folder open completed. folder=" + text + ", reason=" + (reason ?? string.Empty) + ", windowFound=" + (intPtr != IntPtr.Zero));
				return intPtr;
			} catch (Exception exception) {
				_logger.Error ("OpenFolderAndWait failed. folder=" + text, exception);
			}
			return IntPtr.Zero;
		}

		internal bool OpenFolder (string folderPath, string reason)
		{
			string text = _pathCompatibilityService.NormalizePath (folderPath);
			if (string.IsNullOrWhiteSpace (text)) {
				return false;
			}
			try {
				return OpenFolderCore (text, reason);
			} catch (Exception exception) {
				_logger.Error ("OpenFolder failed. folder=" + text, exception);
			}
			return false;
		}

		private bool OpenFolderCore (string normalizedFolderPath, string reason)
		{
			_logger.Info ("Folder open requested. folder=" + normalizedFolderPath + ", reason=" + (reason ?? string.Empty));
			if (Hooks != null && Hooks.StartFolderProcess != null) {
				return Hooks.StartFolderProcess (normalizedFolderPath, reason);
			}
			Process.Start (new ProcessStartInfo {
				FileName = "explorer.exe",
				Arguments = "\"" + normalizedFolderPath + "\"",
				UseShellExecute = true
			});
			return true;
		}

		private IntPtr WaitForFolderWindow (string folderPath)
		{
			DateTime dateTime = DateTime.UtcNow.AddMilliseconds (1500.0);
			while (DateTime.UtcNow < dateTime) {
				IntPtr intPtr = TryFindExplorerWindow (folderPath);
				if (intPtr != IntPtr.Zero) {
					return intPtr;
				}
				Thread.Sleep (100);
			}
			return IntPtr.Zero;
		}

		private IntPtr TryFindExplorerWindow (string folderPath)
		{
			if (Hooks != null && Hooks.TryFindExplorerWindow != null) {
				return Hooks.TryFindExplorerWindow (folderPath);
			}
			object obj = null;
			object obj2 = null;
			try {
				Type typeFromProgID = Type.GetTypeFromProgID ("Shell.Application");
				if (typeFromProgID == null) {
					return IntPtr.Zero;
				}
				obj = Activator.CreateInstance (typeFromProgID);
				dynamic val = obj;
				obj2 = val.Windows ();
				dynamic val2 = obj2;
				int num = Convert.ToInt32 (val2.Count);
				for (int i = 0; i < num; i++) {
					object obj3 = null;
					object obj4 = null;
					object obj5 = null;
					object obj6 = null;
					try {
						obj3 = val2.Item (i);
						if (obj3 == null) {
							continue;
						}
						dynamic val3 = obj3;
						obj4 = val3.Document;
						if (obj4 == null) {
							continue;
						}
						dynamic val4 = obj4;
						obj5 = val4.Folder;
						if (obj5 == null) {
							continue;
						}
						dynamic val5 = obj5;
						obj6 = val5.Self;
						if (obj6 != null) {
							dynamic val6 = obj6;
							string a = _pathCompatibilityService.NormalizePath (Convert.ToString (val6.Path) ?? string.Empty);
							if (string.Equals (a, folderPath, StringComparison.OrdinalIgnoreCase)) {
								return new IntPtr (Convert.ToInt32 (val3.HWND));
							}
						}
					} catch {
					} finally {
						ReleaseComObject (obj6);
						ReleaseComObject (obj5);
						ReleaseComObject (obj4);
						ReleaseComObject (obj3);
					}
				}
			} catch {
			} finally {
				ReleaseComObject (obj2);
				ReleaseComObject (obj);
			}
			return IntPtr.Zero;
		}

		private static void ReleaseComObject (object comObject)
		{
			// Shell COM 参照はこの service 側で寿命を完結させるため完全解放を維持する。
			ComObjectReleaseService.FinalRelease (comObject);
		}
	}
}
