using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
	internal sealed class LocalWorkCopyService : IDisposable
	{
		internal sealed class LocalWorkCopyServiceTestHooks
		{
			internal Func<object, string, bool> IsDocumentOpen { get; set; }

			internal Func<string, bool> FileExistsSafe { get; set; }

			internal Func<string, string, bool> MoveFileSafe { get; set; }

			internal Action<string, string> ShowFinalizeFailureMessage { get; set; }
		}

		private sealed class LocalWorkCopyJob
		{
			internal string LocalPath { get; }

			internal string FinalPath { get; }

			internal string DocumentName { get; }

			internal object WordApplication { get; }

			internal LocalWorkCopyJob (string localPath, string finalPath, string documentName, object wordApplication)
			{
				LocalPath = localPath ?? string.Empty;
				FinalPath = finalPath ?? string.Empty;
				DocumentName = documentName ?? string.Empty;
				WordApplication = wordApplication;
			}
		}

		private const int LocalWorkPollIntervalMs = 5000;

		private readonly PathCompatibilityService _pathCompatibilityService;

		private readonly WordInteropService _wordInteropService;

		private readonly Logger _logger;

		private readonly LocalWorkCopyServiceTestHooks _testHooks;

		private readonly Dictionary<string, LocalWorkCopyJob> _jobs;

		private readonly Timer _pollTimer;

		private bool _disposed;

		internal LocalWorkCopyService (PathCompatibilityService pathCompatibilityService, WordInteropService wordInteropService, Logger logger)
			: this (pathCompatibilityService, wordInteropService, logger, testHooks: null)
		{
		}

		internal LocalWorkCopyService (PathCompatibilityService pathCompatibilityService, WordInteropService wordInteropService, Logger logger, LocalWorkCopyServiceTestHooks testHooks)
		{
			_pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException ("pathCompatibilityService");
			_wordInteropService = wordInteropService ?? throw new ArgumentNullException ("wordInteropService");
			_logger = logger ?? throw new ArgumentNullException ("logger");
			_testHooks = testHooks;
			_jobs = new Dictionary<string, LocalWorkCopyJob> (StringComparer.OrdinalIgnoreCase);
			_pollTimer = new Timer ();
			_pollTimer.Interval = LocalWorkPollIntervalMs;
			_pollTimer.Tick += PollTimer_Tick;
		}

		internal string BuildLocalWorkCopyPath (string finalPath)
		{
			string text = _pathCompatibilityService.NormalizePath (finalPath);
			if (text.Length == 0) {
				return string.Empty;
			}
			string localTempWorkFolder = _pathCompatibilityService.GetLocalTempWorkFolder ("CaseDocTemp");
			if (localTempWorkFolder.Length == 0) {
				return string.Empty;
			}
			string fileNameFromPath = _pathCompatibilityService.GetFileNameFromPath (text);
			if (fileNameFromPath.Length == 0) {
				return string.Empty;
			}
			int num = fileNameFromPath.LastIndexOf ('.');
			string baseName = ((num > 1) ? fileNameFromPath.Substring (0, num) : fileNameFromPath);
			string extension = ((num > 1) ? fileNameFromPath.Substring (num) : ".docx");
			return _pathCompatibilityService.BuildUniquePath (localTempWorkFolder, baseName, extension);
		}

		internal void RegisterLocalWorkCopy (object wordApplication, string localPath, string finalPath)
		{
			string text = _pathCompatibilityService.NormalizePath (localPath);
			string text2 = _pathCompatibilityService.NormalizePath (finalPath);
			if (text.Length != 0 && text2.Length != 0) {
				string key = text.ToLowerInvariant ();
				_jobs[key] = new LocalWorkCopyJob (text, text2, _pathCompatibilityService.GetFileNameFromPath (text2), wordApplication);
				if (!_pollTimer.Enabled) {
					_pollTimer.Start ();
				}
				_logger.Info ("LocalWorkCopyService registered local copy. local=" + text + ", final=" + text2);
			}
		}

		internal bool HasPendingLocalWorkCopies ()
		{
			return _jobs.Count > 0;
		}

		internal string GetPendingLocalWorkCopySummary ()
		{
			if (_jobs.Count == 0) {
				return string.Empty;
			}
			List<string> list = new List<string> ();
			foreach (LocalWorkCopyJob value in _jobs.Values) {
				list.Add ("- " + value.DocumentName);
			}
			return string.Join (Environment.NewLine, list);
		}

		internal void Cancel ()
		{
			_pollTimer.Stop ();
		}

		internal void ExecutePollLocalWorkCopiesForTesting ()
		{
			PollLocalWorkCopies ();
		}

		internal void RaisePollTimerTickForTesting ()
		{
			PollTimer_Tick (this, EventArgs.Empty);
		}

		internal bool IsPollingActiveForTesting ()
		{
			return _pollTimer.Enabled;
		}

		private void PollLocalWorkCopies ()
		{
			if (_jobs.Count == 0) {
				_pollTimer.Stop ();
				return;
			}
			List<string> list = new List<string> ();
			foreach (KeyValuePair<string, LocalWorkCopyJob> job in _jobs) {
				LocalWorkCopyJob value = job.Value;
				if (!IsDocumentOpen (value.WordApplication, value.LocalPath) && FinalizeTrackedWordDoc (value)) {
					list.Add (job.Key);
				}
			}
			foreach (string item in list) {
				_jobs.Remove (item);
			}
			if (_jobs.Count == 0) {
				_pollTimer.Stop ();
			}
		}

		private bool FinalizeTrackedWordDoc (LocalWorkCopyJob job)
		{
			if (job == null) {
				return false;
			}
			if (!FileExistsSafe (job.LocalPath)) {
				return true;
			}
			if (MoveFileSafe (job.LocalPath, job.FinalPath)) {
				_logger.Info ("LocalWorkCopyService finalized local copy. final=" + job.FinalPath);
				return true;
			}
			ShowFinalizeFailureMessage (job);
			return false;
		}

		private void PollTimer_Tick (object sender, EventArgs e)
		{
			try {
				PollLocalWorkCopies ();
			} catch (Exception exception) {
				_logger.Error ("LocalWorkCopyService.PollLocalWorkCopies failed.", exception);
			}
		}

		private bool IsDocumentOpen (object wordApplication, string localPath)
		{
			if (_testHooks != null && _testHooks.IsDocumentOpen != null) {
				return _testHooks.IsDocumentOpen (wordApplication, localPath);
			}
			return _wordInteropService.IsDocumentOpen (wordApplication, localPath);
		}

		private bool FileExistsSafe (string path)
		{
			return (_testHooks != null && _testHooks.FileExistsSafe != null) ? _testHooks.FileExistsSafe (path) : _pathCompatibilityService.FileExistsSafe (path);
		}

		private bool MoveFileSafe (string sourcePath, string destinationPath)
		{
			return (_testHooks != null && _testHooks.MoveFileSafe != null) ? _testHooks.MoveFileSafe (sourcePath, destinationPath) : _pathCompatibilityService.MoveFileSafe (sourcePath, destinationPath);
		}

		private void ShowFinalizeFailureMessage (LocalWorkCopyJob job)
		{
			if (job == null) {
				return;
			}
			if (_testHooks != null && _testHooks.ShowFinalizeFailureMessage != null) {
				_testHooks.ShowFinalizeFailureMessage (job.LocalPath, job.FinalPath);
				return;
			}
			string message = "Could not move the local work copy to the final location."
				+ Environment.NewLine
				+ "Close Word and confirm the destination is available, then try again."
				+ Environment.NewLine
				+ Environment.NewLine
				+ "Local: " + job.LocalPath
				+ Environment.NewLine
				+ "Final: " + job.FinalPath;
			MessageBox.Show (message, "Case System", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
		}

		public void Dispose ()
		{
			if (!_disposed) {
				_disposed = true;
				_pollTimer.Stop ();
				_pollTimer.Tick -= PollTimer_Tick;
				_pollTimer.Dispose ();
			}
		}
	}
}
