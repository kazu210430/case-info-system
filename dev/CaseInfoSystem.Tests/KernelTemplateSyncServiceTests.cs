using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security;
using System.Text;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using CaseInfoSystem.Tests.Fakes;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.Tests
{
	public sealed class KernelTemplateSyncServiceTests
	{
		[Fact]
		public void Execute_WhenPreflightFails_DoesNotRunPublicationSideEffects ()
		{
			using (TestHarness harness = TestHarness.Create (includeDefinedTags: false, createValidTemplate: false, createBaseFile: false)) {
				KernelTemplateSyncResult result = harness.Service.Execute (harness.Context);

				Assert.False (result.Success);
				Assert.Empty (harness.Events);
				Assert.Equal (0, harness.KernelWorkbook.SaveCallCount);
				Assert.False (harness.KernelProperties.ContainsKey ("TASKPANE_MASTER_VERSION"));
			}
		}

		[Fact]
		public void Execute_WhenPublicationSucceeds_RunsSideEffectsInCurrentOrder ()
		{
			using (TestHarness harness = TestHarness.Create (includeDefinedTags: true, createValidTemplate: true, createBaseFile: true)) {
				harness.KernelWorkbook.SaveBehavior = () => harness.Events.Add ("kernel-save");
				harness.BaseWorkbook.SaveBehavior = () => harness.Events.Add ("base-save");

				KernelTemplateSyncResult result = harness.Service.Execute (harness.Context);

				Assert.True (result.Success);
				Assert.Equal (new[] { "master-write", "kernel-version", "kernel-save", "base-save", "invalidate" }, harness.Events);
				Assert.Equal ("1", harness.KernelProperties["TASKPANE_MASTER_VERSION"]);
				Assert.True (string.IsNullOrEmpty (result.BaseSyncError));
				Assert.Equal (1, harness.BaseWorkbook.CloseCallCount);
			}
		}

		[Fact]
		public void Execute_WhenKernelSaveFails_DoesNotInvalidateOrSyncBase ()
		{
			using (TestHarness harness = TestHarness.Create (includeDefinedTags: true, createValidTemplate: true, createBaseFile: false)) {
				harness.KernelWorkbook.SaveBehavior = () =>
				{
					harness.Events.Add ("kernel-save");
					throw new InvalidOperationException ("kernel save failed");
				};

				InvalidOperationException exception = Assert.Throws<InvalidOperationException> (() => harness.Service.Execute (harness.Context));

				Assert.Equal ("kernel save failed", exception.Message);
				Assert.Equal (new[] { "master-write", "kernel-version", "kernel-save" }, harness.Events);
				Assert.Equal (0, harness.BaseWorkbook.SaveCallCount);
				Assert.Equal (0, harness.BaseWorkbook.CloseCallCount);
			}
		}

		[Fact]
		public void Execute_WhenBaseSnapshotSaveFails_ReturnsSuccessWithWarningAndInvalidates ()
		{
			using (TestHarness harness = TestHarness.Create (includeDefinedTags: true, createValidTemplate: true, createBaseFile: true)) {
				harness.KernelWorkbook.SaveBehavior = () => harness.Events.Add ("kernel-save");
				harness.BaseWorkbook.SaveBehavior = () =>
				{
					harness.Events.Add ("base-save");
					throw new InvalidOperationException ("base save failed");
				};

				KernelTemplateSyncResult result = harness.Service.Execute (harness.Context);

				Assert.True (result.Success);
				Assert.Equal (new[] { "master-write", "kernel-version", "kernel-save", "base-save", "invalidate" }, harness.Events);
				Assert.False (string.IsNullOrWhiteSpace (result.BaseSyncError));
				Assert.Contains ("Base", result.Message);
				Assert.Equal (1, harness.BaseWorkbook.CloseCallCount);
			}
		}

		[Fact]
		public void SaveSnapshotToBaseWorkbook_WhenSnapshotIsEmpty_WritesCountZeroAndBaseVersion ()
		{
			using (TestHarness harness = TestHarness.Create (includeDefinedTags: false, createValidTemplate: false, createBaseFile: false)) {
				InvokeSaveSnapshotToBaseWorkbook (harness, string.Empty, 7);

				Assert.Equal ("0", harness.BaseProperties["TASKPANE_BASE_SNAPSHOT_COUNT"]);
				Assert.Equal ("7", harness.BaseProperties["TASKPANE_BASE_MASTER_VERSION"]);
				Assert.False (harness.BaseProperties.ContainsKey ("TASKPANE_MASTER_VERSION"));
			}
		}

		[Fact]
		public void SaveSnapshotToBaseWorkbook_WhenSnapshotExceedsChunkSize_SplitsIntoMultipleChunks ()
		{
			using (TestHarness harness = TestHarness.Create (includeDefinedTags: false, createValidTemplate: false, createBaseFile: false)) {
				string snapshotText = new string ('A', 240) + new string ('B', 240) + "C";

				InvokeSaveSnapshotToBaseWorkbook (harness, snapshotText, 11);

				Assert.Equal ("3", harness.BaseProperties["TASKPANE_BASE_SNAPSHOT_COUNT"]);
				Assert.Equal ("11", harness.BaseProperties["TASKPANE_BASE_MASTER_VERSION"]);
				Assert.Equal ("11", harness.BaseProperties["TASKPANE_MASTER_VERSION"]);
				Assert.Equal (new string ('A', 240), harness.BaseProperties["TASKPANE_BASE_SNAPSHOT_01"]);
				Assert.Equal (new string ('B', 240), harness.BaseProperties["TASKPANE_BASE_SNAPSHOT_02"]);
				Assert.Equal ("C", harness.BaseProperties["TASKPANE_BASE_SNAPSHOT_03"]);
			}
		}

		[Fact]
		public void SaveSnapshotToBaseWorkbook_WhenNewSnapshotIsShorter_ClearsStaleChunks ()
		{
			using (TestHarness harness = TestHarness.Create (includeDefinedTags: false, createValidTemplate: false, createBaseFile: false)) {
				InvokeSaveSnapshotToBaseWorkbook (harness, new string ('X', 481), 3);
				InvokeSaveSnapshotToBaseWorkbook (harness, "short", 4);

				Assert.Equal ("1", harness.BaseProperties["TASKPANE_BASE_SNAPSHOT_COUNT"]);
				Assert.Equal ("4", harness.BaseProperties["TASKPANE_BASE_MASTER_VERSION"]);
				Assert.Equal ("4", harness.BaseProperties["TASKPANE_MASTER_VERSION"]);
				Assert.Equal ("short", harness.BaseProperties["TASKPANE_BASE_SNAPSHOT_01"]);
				Assert.Equal (string.Empty, harness.BaseProperties["TASKPANE_BASE_SNAPSHOT_02"]);
				Assert.Equal (string.Empty, harness.BaseProperties["TASKPANE_BASE_SNAPSHOT_03"]);
			}
		}

		[Fact]
		public void WriteToMasterList_WhenTemplatesContainInvalidOrOutOfRangeKeys_WritesOnlyValidRows ()
		{
			using (TestHarness harness = TestHarness.Create (includeDefinedTags: false, createValidTemplate: false, createBaseFile: false)) {
				object[,] capturedValues = null;
				string capturedAddress = null;
				harness.AccountingWorkbookService.OnWriteRangeValues = (worksheet, address, values) =>
				{
					capturedAddress = address;
					capturedValues = values;
				};

				int updatedCount = InvokeWriteToMasterList (harness, new[] {
					new TemplateRegistrationValidationEntry { Key = "02", FileName = "02_Second.docx", DisplayName = "Second" },
					new TemplateRegistrationValidationEntry { Key = "100", FileName = "100_OutOfRange.docx", DisplayName = "OutOfRange" },
					new TemplateRegistrationValidationEntry { Key = "abc", FileName = "abc_Invalid.docx", DisplayName = "Invalid" },
					new TemplateRegistrationValidationEntry { Key = "0", FileName = "00_Zero.docx", DisplayName = "Zero" },
					null
				});

				Assert.Equal (1, updatedCount);
				Assert.Equal ("$A$3:$C$101", capturedAddress);
				Assert.NotNull (capturedValues);
				Assert.Equal (updatedCount, CountWrittenRows (capturedValues));
				Assert.Null (capturedValues[0, 0]);
				Assert.Equal ("02", capturedValues[1, 0]);
				Assert.Equal ("02_Second.docx", capturedValues[1, 1]);
				Assert.Equal ("Second", capturedValues[1, 2]);
			}
		}

		[Fact]
		public void WriteToMasterList_WhenDisplayNameIsNull_FallsBackToExtractedDocumentName ()
		{
			using (TestHarness harness = TestHarness.Create (includeDefinedTags: false, createValidTemplate: false, createBaseFile: false)) {
				object[,] capturedValues = null;
				harness.AccountingWorkbookService.OnWriteRangeValues = (worksheet, address, values) => capturedValues = values;

				int updatedCount = InvokeWriteToMasterList (harness, new[] {
					new TemplateRegistrationValidationEntry { Key = "01", FileName = "01_契約書.docx", DisplayName = null }
				});

				Assert.Equal (1, updatedCount);
				Assert.NotNull (capturedValues);
				Assert.Equal ("01", capturedValues[0, 0]);
				Assert.Equal ("01_契約書.docx", capturedValues[0, 1]);
				Assert.Equal ("契約書", capturedValues[0, 2]);
			}
		}

		[Fact]
		public void Execute_WhenMasterSheetStartsProtected_TemporarilyUnprotectsAndRestoresOriginalProtectionState ()
		{
			using (TestHarness harness = TestHarness.Create (includeDefinedTags: true, createValidTemplate: true, createBaseFile: true)) {
				bool wasUnprotectedDuringWrite = false;
				ConfigureMasterSheetAsProtected (harness.MasterSheet);
				harness.AccountingWorkbookService.OnWriteRangeValues = (worksheet, address, values) =>
				{
					harness.Events.Add ("master-write");
					wasUnprotectedDuringWrite = !harness.MasterSheet.ProtectContents
						&& !harness.MasterSheet.ProtectDrawingObjects
						&& !harness.MasterSheet.ProtectScenarios;
					ClearMasterSheetProtectionFlags (harness.MasterSheet);
				};
				harness.KernelWorkbook.SaveBehavior = () => harness.Events.Add ("kernel-save");
				harness.BaseWorkbook.SaveBehavior = () => harness.Events.Add ("base-save");

				KernelTemplateSyncResult result = harness.Service.Execute (harness.Context);

				Assert.True (result.Success);
				Assert.True (wasUnprotectedDuringWrite);
				Assert.Equal (1, harness.MasterSheet.UnprotectCallCount);
				Assert.Equal (1, harness.MasterSheet.ProtectCallCount);
				Assert.True (harness.MasterSheet.ProtectContents);
				Assert.True (harness.MasterSheet.ProtectDrawingObjects);
				Assert.True (harness.MasterSheet.ProtectScenarios);
				Assert.Equal (Excel.XlEnableSelection.xlUnlockedCells, harness.MasterSheet.EnableSelection);
				Assert.True (harness.MasterSheet.Protection.AllowFormattingCells);
				Assert.False (harness.MasterSheet.Protection.AllowFormattingColumns);
				Assert.True (harness.MasterSheet.Protection.AllowFormattingRows);
				Assert.False (harness.MasterSheet.Protection.AllowInsertingColumns);
				Assert.True (harness.MasterSheet.Protection.AllowInsertingRows);
				Assert.False (harness.MasterSheet.Protection.AllowInsertingHyperlinks);
				Assert.True (harness.MasterSheet.Protection.AllowDeletingColumns);
				Assert.False (harness.MasterSheet.Protection.AllowDeletingRows);
				Assert.True (harness.MasterSheet.Protection.AllowSorting);
				Assert.False (harness.MasterSheet.Protection.AllowFiltering);
				Assert.True (harness.MasterSheet.Protection.AllowUsingPivotTables);
			}
		}

		[Fact]
		public void Execute_WhenMasterSheetStartsUnprotected_DoesNotRunProtectionTransitions ()
		{
			using (TestHarness harness = TestHarness.Create (includeDefinedTags: true, createValidTemplate: true, createBaseFile: true)) {
				bool wasUnprotectedDuringWrite = false;
				harness.MasterSheet.EnableSelection = Excel.XlEnableSelection.xlNoRestrictions;
				harness.AccountingWorkbookService.OnWriteRangeValues = (worksheet, address, values) =>
				{
					harness.Events.Add ("master-write");
					wasUnprotectedDuringWrite = !harness.MasterSheet.ProtectContents
						&& !harness.MasterSheet.ProtectDrawingObjects
						&& !harness.MasterSheet.ProtectScenarios;
				};
				harness.KernelWorkbook.SaveBehavior = () => harness.Events.Add ("kernel-save");
				harness.BaseWorkbook.SaveBehavior = () => harness.Events.Add ("base-save");

				KernelTemplateSyncResult result = harness.Service.Execute (harness.Context);

				Assert.True (result.Success);
				Assert.True (wasUnprotectedDuringWrite);
				Assert.Equal (0, harness.MasterSheet.UnprotectCallCount);
				Assert.Equal (0, harness.MasterSheet.ProtectCallCount);
				Assert.False (harness.MasterSheet.ProtectContents);
				Assert.False (harness.MasterSheet.ProtectDrawingObjects);
				Assert.False (harness.MasterSheet.ProtectScenarios);
				Assert.Equal (Excel.XlEnableSelection.xlNoRestrictions, harness.MasterSheet.EnableSelection);
			}
		}

		[Fact]
		public void Execute_WhenMasterSheetRestoreFails_SwallowsRestoreFailure ()
		{
			using (TestHarness harness = TestHarness.Create (includeDefinedTags: true, createValidTemplate: true, createBaseFile: true)) {
				KernelTemplateSyncResult result = null;
				ConfigureMasterSheetAsProtected (harness.MasterSheet);
				harness.MasterSheet.ProtectBehavior = () => throw new InvalidOperationException ("restore failed");
				harness.KernelWorkbook.SaveBehavior = () => harness.Events.Add ("kernel-save");
				harness.BaseWorkbook.SaveBehavior = () => harness.Events.Add ("base-save");

				Exception exception = Record.Exception (() => result = harness.Service.Execute (harness.Context));

				Assert.Null (exception);
				Assert.NotNull (result);
				Assert.True (result.Success);
				Assert.Equal (new[] { "master-write", "kernel-version", "kernel-save", "base-save", "invalidate" }, harness.Events);
				Assert.Equal (1, harness.MasterSheet.UnprotectCallCount);
				Assert.Equal (1, harness.MasterSheet.ProtectCallCount);
				Assert.False (harness.MasterSheet.ProtectContents);
			}
		}

		private static void ConfigureMasterSheetAsProtected (Excel.Worksheet worksheet)
		{
			worksheet.ProtectContents = true;
			worksheet.ProtectDrawingObjects = true;
			worksheet.ProtectScenarios = true;
			worksheet.EnableSelection = Excel.XlEnableSelection.xlNoSelection;
			worksheet.Protection.AllowFormattingCells = true;
			worksheet.Protection.AllowFormattingColumns = false;
			worksheet.Protection.AllowFormattingRows = true;
			worksheet.Protection.AllowInsertingColumns = false;
			worksheet.Protection.AllowInsertingRows = true;
			worksheet.Protection.AllowInsertingHyperlinks = false;
			worksheet.Protection.AllowDeletingColumns = true;
			worksheet.Protection.AllowDeletingRows = false;
			worksheet.Protection.AllowSorting = true;
			worksheet.Protection.AllowFiltering = false;
			worksheet.Protection.AllowUsingPivotTables = true;
		}

		private static void ClearMasterSheetProtectionFlags (Excel.Worksheet worksheet)
		{
			worksheet.Protection.AllowFormattingCells = false;
			worksheet.Protection.AllowFormattingColumns = false;
			worksheet.Protection.AllowFormattingRows = false;
			worksheet.Protection.AllowInsertingColumns = false;
			worksheet.Protection.AllowInsertingRows = false;
			worksheet.Protection.AllowInsertingHyperlinks = false;
			worksheet.Protection.AllowDeletingColumns = false;
			worksheet.Protection.AllowDeletingRows = false;
			worksheet.Protection.AllowSorting = false;
			worksheet.Protection.AllowFiltering = false;
			worksheet.Protection.AllowUsingPivotTables = false;
		}

		private static void InvokeSaveSnapshotToBaseWorkbook (TestHarness harness, string snapshotText, int masterVersion)
		{
			FieldInfo publicationExecutorField = typeof (KernelTemplateSyncService).GetField ("_publicationExecutor", BindingFlags.Instance | BindingFlags.NonPublic);
			Assert.NotNull (publicationExecutorField);
			object publicationExecutor = publicationExecutorField.GetValue (harness.Service);
			Assert.NotNull (publicationExecutor);

			MethodInfo saveSnapshotMethod = publicationExecutor.GetType ().GetMethod ("SaveSnapshotToBaseWorkbook", BindingFlags.Instance | BindingFlags.NonPublic);
			Assert.NotNull (saveSnapshotMethod);
			saveSnapshotMethod.Invoke (publicationExecutor, new object[] { harness.BaseWorkbook, snapshotText, masterVersion });
		}

		private static int InvokeWriteToMasterList (TestHarness harness, IReadOnlyList<TemplateRegistrationValidationEntry> templates)
		{
			FieldInfo publicationExecutorField = typeof (KernelTemplateSyncService).GetField ("_publicationExecutor", BindingFlags.Instance | BindingFlags.NonPublic);
			Assert.NotNull (publicationExecutorField);
			object publicationExecutor = publicationExecutorField.GetValue (harness.Service);
			Assert.NotNull (publicationExecutor);

			MethodInfo writeToMasterListMethod = publicationExecutor.GetType ().GetMethod ("WriteToMasterList", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
			Assert.NotNull (writeToMasterListMethod);
			object result = writeToMasterListMethod.Invoke (publicationExecutor, new object[] { harness.MasterSheet, templates });
			return result is int updatedCount ? updatedCount : 0;
		}

		private static int CountWrittenRows (object[,] values)
		{
			if (values == null) {
				return 0;
			}
			int num = 0;
			for (int i = 0; i < values.GetLength (0); i++) {
				if (values [i, 0] != null) {
					num++;
				}
			}
			return num;
		}

		private sealed class TestHarness : IDisposable
		{
			private const string SystemRootPropertyName = "SYSTEM_ROOT";

			private const string MasterSheetCodeName = "shMasterList";

			private const string MasterSheetName = "\u96db\u5f62\u4e00\u89a7";

			private const string FieldInventorySheetName = "CaseList_FieldInventory";

			private const string TemplateFolderName = "\u96db\u5f62";

			private const string DefinedTag = "CustomerName";

			private readonly string _systemRoot;

			private readonly List<string> _logMessages;

			private readonly Logger _logger;

			internal Excel.Application Application { get; }

			internal ExcelInteropService ExcelInteropService { get; }

			internal AccountingWorkbookService AccountingWorkbookService { get; }

			internal MasterTemplateCatalogService MasterTemplateCatalogService { get; }

			internal KernelTemplateSyncService Service { get; }

			internal WorkbookContext Context { get; }

			internal Excel.Workbook KernelWorkbook { get; }

			internal Excel.Workbook BaseWorkbook { get; }

			internal Excel.Worksheet MasterSheet { get; }

			internal Dictionary<string, string> KernelProperties { get; }

			internal Dictionary<string, string> BaseProperties { get; }

			internal List<string> Events { get; }

			private TestHarness (bool includeDefinedTags, bool createValidTemplate, bool createBaseFile)
			{
				_systemRoot = CreateTempSystemRoot ();
				_logMessages = new List<string> ();
				_logger = OrchestrationTestSupport.CreateLogger (_logMessages);

				Application = new Excel.Application ();
				ExcelInteropService = new ExcelInteropService ();
				AccountingWorkbookService = new AccountingWorkbookService ();
				MasterTemplateCatalogService = new MasterTemplateCatalogService ();
				Events = new List<string> ();

				KernelProperties = new Dictionary<string, string> (StringComparer.OrdinalIgnoreCase) {
					[SystemRootPropertyName] = _systemRoot
				};
				KernelWorkbook = CreateKernelWorkbook (_systemRoot, KernelProperties);
				MasterSheet = KernelWorkbook.Worksheets [1];
				BaseWorkbook = CreateBaseWorkbook (_systemRoot);
				BaseProperties = (Dictionary<string, string>)BaseWorkbook.CustomDocumentProperties;

				if (includeDefinedTags) {
					AddFieldInventorySheet (KernelWorkbook);
				}

				Application.Workbooks.Add (KernelWorkbook);
				Application.ActiveWorkbook = KernelWorkbook;
				Application.ActiveWindow = KernelWorkbook.Windows.Count > 0 ? KernelWorkbook.Windows [1] : null;

				if (createValidTemplate) {
					CreateWordPackage (Path.Combine (GetTemplateDirectory (_systemRoot), "01_Sample.docx"), BuildDocumentXml (CreateTextControl (DefinedTag)));
				}
				if (createBaseFile) {
					File.WriteAllText (GetBaseWorkbookPath (_systemRoot), "base", Encoding.UTF8);
					Application.Workbooks.OpenBehavior = (filename, updateLinks, readOnly) =>
					{
						BaseWorkbook.FullName = filename ?? BaseWorkbook.FullName;
						BaseWorkbook.Name = Path.GetFileName (BaseWorkbook.FullName);
						BaseWorkbook.Path = Path.GetDirectoryName (BaseWorkbook.FullName) ?? string.Empty;
						return BaseWorkbook;
					};
				}

				ExcelInteropService.OnReadRecordsFromHeaderRow = worksheet =>
				{
					if (worksheet == null || !string.Equals (worksheet.Name, FieldInventorySheetName, StringComparison.OrdinalIgnoreCase)) {
						return Array.Empty<IReadOnlyDictionary<string, string>> ();
					}
					return new IReadOnlyDictionary<string, string>[1] {
						new Dictionary<string, string> (StringComparer.OrdinalIgnoreCase) {
							["ProposedFieldKey"] = DefinedTag,
							["Label"] = "Customer Name"
						}
					};
				};
				ExcelInteropService.OnSetDocumentProperty = (workbook, propertyName, value) =>
				{
					if (ReferenceEquals (workbook, KernelWorkbook) && string.Equals (propertyName, "TASKPANE_MASTER_VERSION", StringComparison.OrdinalIgnoreCase)) {
						Events.Add ("kernel-version");
					}
				};
				AccountingWorkbookService.OnWriteRangeValues = (worksheet, address, values) => Events.Add ("master-write");
				MasterTemplateCatalogService.OnInvalidateCache = workbook => Events.Add ("invalidate");

				PathCompatibilityService pathCompatibilityService = new PathCompatibilityService (_logger);
				KernelWorkbookService kernelWorkbookService = new KernelWorkbookService (
					OrchestrationTestSupport.CreateKernelCaseInteractionState (new List<string> ()),
					_logger,
					new KernelWorkbookService.KernelWorkbookServiceTestHooks ());
				CaseListFieldDefinitionRepository caseListFieldDefinitionRepository = new CaseListFieldDefinitionRepository (ExcelInteropService);
				KernelTemplateSyncPreflightService preflightService = new KernelTemplateSyncPreflightService (
					pathCompatibilityService,
					new WordTemplateRegistrationValidationService (new WordTemplateContentControlInspectionService (), _logger));
				CaseWorkbookLifecycleService caseWorkbookLifecycleService = new CaseWorkbookLifecycleService (
					_logger,
					new CaseWorkbookLifecycleService.CaseWorkbookLifecycleServiceTestHooks {
						GetWorkbookKey = workbook => workbook == null ? string.Empty : workbook.FullName ?? string.Empty
					});
				Service = new KernelTemplateSyncService (
					Application,
					kernelWorkbookService,
					ExcelInteropService,
					AccountingWorkbookService,
					pathCompatibilityService,
					caseListFieldDefinitionRepository,
					preflightService,
					MasterTemplateCatalogService,
					caseWorkbookLifecycleService,
					_logger);
				Context = new WorkbookContext (KernelWorkbook, null, WorkbookRole.Kernel, _systemRoot, KernelWorkbook.FullName, MasterSheetCodeName);
			}

			internal static TestHarness Create (bool includeDefinedTags, bool createValidTemplate, bool createBaseFile)
			{
				return new TestHarness (includeDefinedTags, createValidTemplate, createBaseFile);
			}

			public void Dispose ()
			{
				try {
					if (Directory.Exists (_systemRoot)) {
						Directory.Delete (_systemRoot, recursive: true);
					}
				} catch {
				}
			}

			private static Excel.Workbook CreateKernelWorkbook (string systemRoot, Dictionary<string, string> properties)
			{
				Excel.Workbook workbook = new Excel.Workbook {
					Name = WorkbookFileNameResolver.BuildKernelWorkbookName (".xlsm"),
					FullName = Path.Combine (systemRoot, WorkbookFileNameResolver.BuildKernelWorkbookName (".xlsm")),
					Path = systemRoot,
					CustomDocumentProperties = properties
				};
				Excel.Worksheet masterSheet = new Excel.Worksheet {
					CodeName = MasterSheetCodeName,
					Name = MasterSheetName,
					Parent = workbook
				};
				workbook.Worksheets.Add (masterSheet);
				workbook.ActiveSheet = masterSheet;
				return workbook;
			}

			private static Excel.Workbook CreateBaseWorkbook (string systemRoot)
			{
				return new Excel.Workbook {
					Name = WorkbookFileNameResolver.BuildBaseWorkbookName (".xlsm"),
					FullName = GetBaseWorkbookPath (systemRoot),
					Path = systemRoot,
					CustomDocumentProperties = new Dictionary<string, string> (StringComparer.OrdinalIgnoreCase)
				};
			}

			private static void AddFieldInventorySheet (Excel.Workbook workbook)
			{
				workbook.Worksheets.Add (new Excel.Worksheet {
					Name = FieldInventorySheetName,
					Parent = workbook
				});
			}

			private static string CreateTempSystemRoot ()
			{
				string path = Path.Combine (Path.GetTempPath (), "CaseInfoSystem.KernelTemplateSync." + Guid.NewGuid ().ToString ("N"));
				Directory.CreateDirectory (path);
				Directory.CreateDirectory (GetTemplateDirectory (path));
				return path;
			}

			private static string GetTemplateDirectory (string systemRoot)
			{
				return Path.Combine (systemRoot, TemplateFolderName);
			}

			private static string GetBaseWorkbookPath (string systemRoot)
			{
				return Path.Combine (systemRoot, WorkbookFileNameResolver.BuildBaseWorkbookName (".xlsm"));
			}

			private static void CreateWordPackage (string fullPath, string documentXml)
			{
				using (ZipArchive zipArchive = ZipFile.Open (fullPath, ZipArchiveMode.Create)) {
					ZipArchiveEntry entry = zipArchive.CreateEntry ("word/document.xml");
					using (StreamWriter writer = new StreamWriter (entry.Open (), new UTF8Encoding (encoderShouldEmitUTF8Identifier: false))) {
						writer.Write (documentXml);
					}
				}
			}

			private static string BuildDocumentXml (params string[] contentControls)
			{
				return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
					+ "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
					+ "<w:body>"
					+ string.Concat (contentControls ?? Array.Empty<string> ())
					+ "<w:sectPr/>"
					+ "</w:body>"
					+ "</w:document>";
			}

			private static string CreateTextControl (string tag)
			{
				return "<w:sdt><w:sdtPr>" + BuildTagElement (tag) + "<w:text/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>sample</w:t></w:r></w:p></w:sdtContent></w:sdt>";
			}

			private static string BuildTagElement (string tag)
			{
				return tag == null ? string.Empty : "<w:tag w:val=\"" + SecurityElement.Escape (tag) + "\"/>";
			}
		}
	}
}
