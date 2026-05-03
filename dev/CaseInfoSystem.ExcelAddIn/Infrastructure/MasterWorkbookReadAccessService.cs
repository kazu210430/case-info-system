using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    internal enum MasterWorkbookPathResolutionMode
    {
        MasterTemplateCatalog,
        TaskPaneSnapshotBuilder
    }

    internal enum MasterWorkbookOpenSearchMode
    {
        StrictFullPathOnly,
        FullPathOrFileName
    }

    internal sealed class MasterWorkbookReadAccessResult
    {
        internal MasterWorkbookReadAccessResult(string resolvedMasterPath, Excel.Workbook workbook, bool workbookWasAlreadyOpen)
        {
            ResolvedMasterPath = resolvedMasterPath ?? string.Empty;
            Workbook = workbook;
            WorkbookWasAlreadyOpen = workbookWasAlreadyOpen;
        }

        internal string ResolvedMasterPath { get; }

        internal Excel.Workbook Workbook { get; }

        internal bool WorkbookWasAlreadyOpen { get; }

        internal void CloseIfOwned()
        {
            if (WorkbookWasAlreadyOpen || Workbook == null)
            {
                return;
            }

            Workbook.Close(false, Type.Missing, Type.Missing);
        }
    }

    internal sealed class MasterWorkbookReadAccessService
    {
        private const string SystemRootPropertyName = "SYSTEM_ROOT";

        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly PathCompatibilityService _pathCompatibilityService;

        internal MasterWorkbookReadAccessService(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
        }

        internal string ResolveMasterPath(Excel.Workbook workbook, MasterWorkbookPathResolutionMode resolutionMode)
        {
            return resolutionMode == MasterWorkbookPathResolutionMode.TaskPaneSnapshotBuilder
                ? ResolveMasterPathForSnapshotBuilder(workbook)
                : ResolveMasterPathForMasterTemplateCatalog(workbook);
        }

        internal Excel.Workbook FindOpenMasterWorkbook(string resolvedMasterPath, MasterWorkbookOpenSearchMode searchMode)
        {
            if (string.IsNullOrWhiteSpace(resolvedMasterPath))
            {
                return null;
            }

            Excel.Workbook workbook = _excelInteropService.FindOpenWorkbook(resolvedMasterPath);
            if (workbook != null)
            {
                return workbook;
            }

            if (searchMode != MasterWorkbookOpenSearchMode.FullPathOrFileName)
            {
                return null;
            }

            string fileNameFromPath = _pathCompatibilityService.GetFileNameFromPath(resolvedMasterPath);
            if (string.IsNullOrWhiteSpace(fileNameFromPath))
            {
                return null;
            }

            foreach (Excel.Workbook current in _application.Workbooks)
            {
                string workbookName = _excelInteropService.GetWorkbookName(current);
                if (string.Equals(workbookName, fileNameFromPath, StringComparison.OrdinalIgnoreCase))
                {
                    return current;
                }
            }

            return null;
        }

        internal MasterWorkbookReadAccessResult OpenReadOnly(
            string resolvedMasterPath,
            MasterWorkbookOpenSearchMode searchMode,
            Func<string, Exception> missingWorkbookExceptionFactory)
        {
            if (string.IsNullOrWhiteSpace(resolvedMasterPath))
            {
                throw new InvalidOperationException("Master workbook path could not be resolved.");
            }

            Excel.Workbook openWorkbook = FindOpenMasterWorkbook(resolvedMasterPath, searchMode);
            if (openWorkbook != null)
            {
                return new MasterWorkbookReadAccessResult(resolvedMasterPath, openWorkbook, workbookWasAlreadyOpen: true);
            }

            if (!_pathCompatibilityService.FileExistsSafe(resolvedMasterPath))
            {
                throw (missingWorkbookExceptionFactory ?? CreateDefaultMissingWorkbookExceptionFactory())(resolvedMasterPath);
            }

            bool previousEnableEvents = _application.EnableEvents;
            try
            {
                _application.EnableEvents = false;
                Excel.Workbook workbook = _application.Workbooks.Open(
                    resolvedMasterPath,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    IgnoreReadOnlyRecommended: true,
                    AddToMru: false);
                HideWorkbookWindows(workbook);
                return new MasterWorkbookReadAccessResult(resolvedMasterPath, workbook, workbookWasAlreadyOpen: false);
            }
            finally
            {
                _application.EnableEvents = previousEnableEvents;
            }
        }

        private string ResolveMasterPathForMasterTemplateCatalog(Excel.Workbook workbook)
        {
            string systemRoot = ResolveSystemRootForMasterTemplateCatalog(workbook);
            return systemRoot.Length == 0
                ? string.Empty
                : WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(systemRoot, _pathCompatibilityService);
        }

        private string ResolveMasterPathForSnapshotBuilder(Excel.Workbook workbook)
        {
            string systemRoot = _pathCompatibilityService.NormalizePath(_excelInteropService.TryGetDocumentProperty(workbook, SystemRootPropertyName));
            if (systemRoot.Length > 0)
            {
                return WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(systemRoot, _pathCompatibilityService);
            }

            string workbookPath = _pathCompatibilityService.NormalizePath(_excelInteropService.GetWorkbookPath(workbook));
            if (workbookPath.Length == 0)
            {
                return string.Empty;
            }

            string workbookCandidate = WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(workbookPath, _pathCompatibilityService);
            if (_pathCompatibilityService.FileExistsSafe(workbookCandidate))
            {
                return workbookCandidate;
            }

            string parentPath = _pathCompatibilityService.NormalizePath(_pathCompatibilityService.GetParentFolderPath(workbookPath));
            return parentPath.Length == 0
                ? string.Empty
                : WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(parentPath, _pathCompatibilityService);
        }

        private string ResolveSystemRootForMasterTemplateCatalog(Excel.Workbook workbook)
        {
            string systemRoot = _pathCompatibilityService.NormalizePath(_excelInteropService.TryGetDocumentProperty(workbook, SystemRootPropertyName));
            if (HasExistingMasterWorkbook(systemRoot))
            {
                return systemRoot;
            }

            string workbookPath = _pathCompatibilityService.NormalizePath(_excelInteropService.GetWorkbookPath(workbook));
            if (HasExistingMasterWorkbook(workbookPath))
            {
                return workbookPath;
            }

            string parentPath = _pathCompatibilityService.NormalizePath(_pathCompatibilityService.GetParentFolderPath(workbookPath));
            if (HasExistingMasterWorkbook(parentPath))
            {
                return parentPath;
            }

            return workbookPath;
        }

        private bool HasExistingMasterWorkbook(string systemRoot)
        {
            if (string.IsNullOrWhiteSpace(systemRoot))
            {
                return false;
            }

            return _pathCompatibilityService.FileExistsSafe(
                WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(systemRoot, _pathCompatibilityService));
        }

        private static void HideWorkbookWindows(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return;
            }

            Excel.Windows windows = null;
            try
            {
                windows = workbook.Windows;
                int windowCount = windows == null ? 0 : windows.Count;
                for (int index = 1; index <= windowCount; index++)
                {
                    Excel.Window window = null;
                    try
                    {
                        window = windows[index];
                        if (window != null)
                        {
                            window.Visible = false;
                        }
                    }
                    finally
                    {
                        if (window != null && Marshal.IsComObject(window))
                        {
                            ComObjectReleaseService.Release(window);
                        }
                    }
                }
            }
            catch
            {
                // 非表示化に失敗しても、Master 読み取り自体は継続する。
            }
            finally
            {
                if (windows != null && Marshal.IsComObject(windows))
                {
                    ComObjectReleaseService.Release(windows);
                }
            }
        }

        private static Func<string, Exception> CreateDefaultMissingWorkbookExceptionFactory()
        {
            return resolvedMasterPath => new InvalidOperationException("Masterブックが見つかりません。 path=" + resolvedMasterPath);
        }
    }
}
