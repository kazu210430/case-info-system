using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public class FolderWindowServiceTests
	{
		[Fact]
		public void ConfirmFolderWindow_WhenExplorerWindowIsFoundImmediately_DoesNotStartProcess ()
		{
			List<string> logs = new List<string> ();
			FolderWindowService service = CreateService (logs);
			bool startProcessCalled = false;
			int sleepCalls = 0;
			int tryFindCalls = 0;
			IntPtr expectedWindow = new IntPtr (1234);

			service.Hooks = new FolderWindowService.TestHooks {
				StartFolderProcess = (folderPath, reason) => {
					startProcessCalled = true;
					return true;
				},
				TryFindExplorerWindow = folderPath => {
					tryFindCalls++;
					return expectedWindow;
				},
				Sleep = milliseconds => sleepCalls++
			};

			IntPtr actualWindow = service.ConfirmFolderWindow (BuildTempFolderPath (), "FolderWindowServiceTests.Immediate", 300);

			Assert.Equal (expectedWindow, actualWindow);
			Assert.False (startProcessCalled);
			Assert.Equal (1, tryFindCalls);
			Assert.Equal (0, sleepCalls);
			Assert.Contains (logs, log => log.Contains ("ConfirmFolderWindow probe completed"));
			Assert.Contains (logs, log => log.Contains ("tryFindElapsedMs="));
		}

		[Fact]
		public void ConfirmFolderWindow_WhenExplorerWindowIsNotFound_ReturnsZeroWithoutThrowing ()
		{
			FolderWindowService service = CreateService ();
			int tryFindCalls = 0;

			service.Hooks = new FolderWindowService.TestHooks {
				TryFindExplorerWindow = folderPath => {
					tryFindCalls++;
					return IntPtr.Zero;
				}
			};

			IntPtr actualWindow = service.ConfirmFolderWindow (BuildTempFolderPath (), "FolderWindowServiceTests.NotFound", 0);

			Assert.Equal (IntPtr.Zero, actualWindow);
			Assert.Equal (1, tryFindCalls);
		}

		[Fact]
		public void ConfirmFolderWindow_WhenTimeoutExpires_DoesNotBlockLongAndDoesNotStartProcess ()
		{
			FolderWindowService service = CreateService ();
			bool startProcessCalled = false;
			int tryFindCalls = 0;
			DateTime currentUtc = new DateTime (2026, 4, 24, 0, 0, 0, DateTimeKind.Utc);
			Stopwatch stopwatch = Stopwatch.StartNew ();

			service.Hooks = new FolderWindowService.TestHooks {
				StartFolderProcess = (folderPath, reason) => {
					startProcessCalled = true;
					return true;
				},
				TryFindExplorerWindow = folderPath => {
					tryFindCalls++;
					return IntPtr.Zero;
				},
				Sleep = milliseconds => currentUtc = currentUtc.AddMilliseconds (milliseconds),
				UtcNow = () => currentUtc
			};

			IntPtr actualWindow = service.ConfirmFolderWindow (BuildTempFolderPath (), "FolderWindowServiceTests.Timeout", 300);
			stopwatch.Stop ();

			Assert.Equal (IntPtr.Zero, actualWindow);
			Assert.False (startProcessCalled);
			Assert.InRange (tryFindCalls, 2, 4);
			Assert.True (stopwatch.ElapsedMilliseconds < 200, "ConfirmFolderWindow should not block when test hooks eliminate real sleeping.");
		}

		private static FolderWindowService CreateService (List<string> logs = null)
		{
			return new FolderWindowService (
				new PathCompatibilityService (),
				new Logger (message => {
					if (logs != null) {
						logs.Add (message);
					}
				}));
		}

		private static string BuildTempFolderPath ()
		{
			return Path.Combine (Path.GetTempPath (), "CaseInfoSystem.FolderWindowServiceTests", Guid.NewGuid ().ToString ("N"));
		}
	}
}
