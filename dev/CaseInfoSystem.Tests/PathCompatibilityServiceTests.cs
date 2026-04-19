using System;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Xunit;

namespace CaseInfoSystem.Tests
{
	[CollectionDefinition ("PathCompatibilityServiceEnvironment", DisableParallelization = true)]
	public sealed class PathCompatibilityServiceEnvironmentCollection
	{
	}

	[Collection ("PathCompatibilityServiceEnvironment")]
	public class PathCompatibilityServiceTests
	{
		[Fact]
		public void ResolveToExistingLocalPath_WhenDocsLiveUrlUsesAsciiSegments_ReturnsExistingLocalPath ()
		{
			using TestEnvironmentScope scope = new TestEnvironmentScope ();
			PathCompatibilityService service = new PathCompatibilityService ();
			string expectedPath = Path.Combine (scope.OneDriveRoot, "Docs", "inside.docx");
			Directory.CreateDirectory (Path.GetDirectoryName (expectedPath));
			File.WriteAllText (expectedPath, "test");

			string resolvedPath = service.ResolveToExistingLocalPath ("https://d.docs.live.net/cid123/Docs/inside.docx");

			Assert.Equal (service.NormalizePath (expectedPath), resolvedPath);
		}

		[Fact]
		public void BuildSafeSavePath_WhenDocsLiveUrlUsesJapanesePercentEncoding_ReturnsEmptyStringEvenWhenFolderExists ()
		{
			using TestEnvironmentScope scope = new TestEnvironmentScope ();
			PathCompatibilityService service = new PathCompatibilityService ();
			string existingFolder = Path.Combine (scope.OneDriveRoot, "文書");
			Directory.CreateDirectory (existingFolder);

			string safeSavePath = service.BuildSafeSavePath ("https://d.docs.live.net/cid123/%E6%96%87%E6%9B%B8/%E5%A5%91%E7%B4%84%E6%9B%B8.docx");

			Assert.Equal (string.Empty, safeSavePath);
		}

		[Fact]
		public void ResolveToExistingLocalPath_WhenSharePointUrlUsesJapanesePercentEncoding_ReturnsEmptyStringEvenWhenLocalPathExists ()
		{
			using TestEnvironmentScope scope = new TestEnvironmentScope ();
			PathCompatibilityService service = new PathCompatibilityService ();
			string expectedPath = Path.Combine (scope.OneDriveRoot, "Shared Documents", "文書", "契約書.docx");
			Directory.CreateDirectory (Path.GetDirectoryName (expectedPath));
			File.WriteAllText (expectedPath, "test");

			string resolvedPath = service.ResolveToExistingLocalPath ("https://contoso.sharepoint.com/sites/test/Shared%20Documents/%E6%96%87%E6%9B%B8/%E5%A5%91%E7%B4%84%E6%9B%B8.docx");

			Assert.Equal (string.Empty, resolvedPath);
		}

		[Fact]
		public void MoveFileSafe_WhenDestinationExists_ReplacesFileWithoutDeletingDestinationFirst ()
		{
			using TestEnvironmentScope scope = new TestEnvironmentScope ();
			PathCompatibilityService service = new PathCompatibilityService ();
			string sourceFolder = Path.Combine (scope.TempRoot, "source");
			string destinationFolder = Path.Combine (scope.TempRoot, "destination");
			Directory.CreateDirectory (sourceFolder);
			Directory.CreateDirectory (destinationFolder);
			string sourcePath = Path.Combine (sourceFolder, "draft.docx");
			string destinationPath = Path.Combine (destinationFolder, "result.docx");
			File.WriteAllText (sourcePath, "new");
			File.WriteAllText (destinationPath, "old");

			bool moved = service.MoveFileSafe (sourcePath, destinationPath);

			Assert.True (moved);
			Assert.False (File.Exists (sourcePath));
			Assert.Equal ("new", File.ReadAllText (destinationPath));
			Assert.DoesNotContain (Directory.GetFiles (destinationFolder), path => path != destinationPath);
		}

		[Fact]
		public void MoveFileSafe_WhenDestinationReplacementFails_LeavesExistingFileUntouched ()
		{
			using TestEnvironmentScope scope = new TestEnvironmentScope ();
			PathCompatibilityService service = new PathCompatibilityService ();
			string sourceFolder = Path.Combine (scope.TempRoot, "source");
			string destinationFolder = Path.Combine (scope.TempRoot, "destination");
			Directory.CreateDirectory (sourceFolder);
			Directory.CreateDirectory (destinationFolder);
			string sourcePath = Path.Combine (sourceFolder, "draft.docx");
			string destinationPath = Path.Combine (destinationFolder, "result.docx");
			File.WriteAllText (sourcePath, "new");
			File.WriteAllText (destinationPath, "old");

			bool moved;
			using (FileStream destinationLock = new FileStream (destinationPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
			{
				moved = service.MoveFileSafe (sourcePath, destinationPath);
			}

			Assert.False (moved);
			Assert.True (File.Exists (sourcePath));
			Assert.Equal ("old", File.ReadAllText (destinationPath));
			Assert.DoesNotContain (Directory.GetFiles (destinationFolder), path => path != destinationPath);
		}

		private sealed class TestEnvironmentScope : IDisposable
		{
			private readonly string _originalOneDrive;
			private readonly string _originalOneDriveCommercial;
			private readonly string _originalOneDriveConsumer;
			private readonly string _tempRoot;

			internal TestEnvironmentScope ()
			{
				_originalOneDrive = Environment.GetEnvironmentVariable ("OneDrive");
				_originalOneDriveCommercial = Environment.GetEnvironmentVariable ("OneDriveCommercial");
				_originalOneDriveConsumer = Environment.GetEnvironmentVariable ("OneDriveConsumer");
				_tempRoot = Path.Combine (Path.GetTempPath (), "CaseInfoSystem.PathCompatibilityTests." + Guid.NewGuid ().ToString ("N"));
				Directory.CreateDirectory (_tempRoot);
				OneDriveRoot = Path.Combine (_tempRoot, "OneDrive");
				Directory.CreateDirectory (OneDriveRoot);
				Environment.SetEnvironmentVariable ("OneDrive", OneDriveRoot);
				Environment.SetEnvironmentVariable ("OneDriveCommercial", OneDriveRoot);
				Environment.SetEnvironmentVariable ("OneDriveConsumer", OneDriveRoot);
			}

			internal string OneDriveRoot { get; }

			internal string TempRoot => _tempRoot;

			public void Dispose ()
			{
				Environment.SetEnvironmentVariable ("OneDrive", _originalOneDrive);
				Environment.SetEnvironmentVariable ("OneDriveCommercial", _originalOneDriveCommercial);
				Environment.SetEnvironmentVariable ("OneDriveConsumer", _originalOneDriveConsumer);
				try {
					if (Directory.Exists (_tempRoot)) {
						Directory.Delete (_tempRoot, recursive: true);
					}
				} catch {
				}
			}
		}
	}
}
