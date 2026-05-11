using System;
using System.Collections.Generic;
using System.IO;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal interface IKernelUserDataReflectionService
    {
        void ReflectAll(WorkbookContext context);

        void ReflectToAccountingSetOnly(WorkbookContext context);

        void ReflectToBaseHomeOnly(WorkbookContext context);
    }

    /// <summary>
    /// Kernel のユーザー情報を Base HOME と会計セットへ反映するサービス。
    /// shUserData の値を読み取り、転記先ブックへ静かに書き戻す。
    /// </summary>
    internal sealed class KernelUserDataReflectionService : IKernelUserDataReflectionService
    {
        private const string CaseHomeSheetCodeName = "shHOME";
        private const string CaseHomeSheetName = "ホーム";
        private const string UnprotectPassword = "";
        private const string UserDataPostalCodeKey = FieldKeyRenameMap.CurrentPostalCodeKey;
        private const string UserDataAddressKey = FieldKeyRenameMap.CurrentAddressKey;
        private const string UserDataOfficeNameKey = FieldKeyRenameMap.CurrentOfficeNameKey;
        private const string UserDataPhoneKey = FieldKeyRenameMap.CurrentPhoneKey;
        private const string UserDataAccountingNameRow1Key = "銀行・支店";
        private const string UserDataAccountingNameRow2Key = "口座番号・名義";
        private const string HiddenExcelCleanupCompleted = "HiddenExcelCleanupCompleted";
        private const string HiddenExcelCleanupDegraded = "HiddenExcelCleanupDegraded";
        private const string HiddenExcelCleanupNotRequired = "HiddenExcelCleanupNotRequired";
        private const string IsolatedAppReleased = "IsolatedAppReleased";
        private const string IsolatedAppReleaseDegraded = "IsolatedAppReleaseDegraded";
        private const string IsolatedAppReleaseNotRequired = "IsolatedAppReleaseNotRequired";
        private const string IsolatedAppReleaseFailed = "IsolatedAppReleaseFailed";
        private const string ApplicationKindIsolated = "isolated";
        private const string ApplicationKindSharedCurrent = "shared-current";
        private const string ApplicationKindRetainedHiddenAppCache = "retained-hidden-app-cache";
        private const string ApplicationLifetimeOwnerKernelUserDataReflectionService = "KernelUserDataReflectionService";
        private const string ApplicationLifetimeOwnerUserOrExcelHost = "user-or-excel-host";

        private readonly KernelWorkbookService _kernelWorkbookService;
        private readonly ExcelInteropService _excelInteropService;
        private readonly AccountingTemplateResolver _accountingTemplateResolver;
        private readonly AccountingWorkbookService _accountingWorkbookService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly UserDataBaseMappingRepository _userDataBaseMappingRepository;
        private readonly Logger _logger;

        internal KernelUserDataReflectionService(
            KernelWorkbookService kernelWorkbookService,
            ExcelInteropService excelInteropService,
            AccountingTemplateResolver accountingTemplateResolver,
            AccountingWorkbookService accountingWorkbookService,
            PathCompatibilityService pathCompatibilityService,
            UserDataBaseMappingRepository userDataBaseMappingRepository,
            Logger logger)
        {
            _kernelWorkbookService = kernelWorkbookService ?? throw new ArgumentNullException(nameof(kernelWorkbookService));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _accountingTemplateResolver = accountingTemplateResolver ?? throw new ArgumentNullException(nameof(accountingTemplateResolver));
            _accountingWorkbookService = accountingWorkbookService ?? throw new ArgumentNullException(nameof(accountingWorkbookService));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _userDataBaseMappingRepository = userDataBaseMappingRepository ?? throw new ArgumentNullException(nameof(userDataBaseMappingRepository));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// メソッド: ユーザー情報を会計セットと Base HOME の両方へ反映する。
        /// 副作用: 必要に応じて対象ブックを開き、セル値と DocProp を更新する。
        /// </summary>
        internal void ReflectAll(WorkbookContext context)
        {
            ExecuteReflection(
                context,
                (kernelWorkbook, systemRoot, userDataWorksheet, snapshot) =>
                {
                    ReflectToAccountingSet(kernelWorkbook, snapshot);
                    ReflectToBaseHome(systemRoot, kernelWorkbook, userDataWorksheet, snapshot);
                    _logger.Info("Kernel user data reflected to accounting set and base home.");
                });
        }

        void IKernelUserDataReflectionService.ReflectAll(WorkbookContext context)
        {
            ReflectAll(context);
        }

        internal void ReflectToAccountingSetOnly(WorkbookContext context)
        {
            ExecuteReflection(
                context,
                (kernelWorkbook, systemRoot, userDataWorksheet, snapshot) =>
                {
                    ReflectToAccountingSet(kernelWorkbook, snapshot);
                    _logger.Info("Kernel user data reflected to accounting set.");
                });
        }

        void IKernelUserDataReflectionService.ReflectToAccountingSetOnly(WorkbookContext context)
        {
            ReflectToAccountingSetOnly(context);
        }

        internal void ReflectToBaseHomeOnly(WorkbookContext context)
        {
            ExecuteReflection(
                context,
                (kernelWorkbook, systemRoot, userDataWorksheet, snapshot) =>
                {
                    ReflectToBaseHome(systemRoot, kernelWorkbook, userDataWorksheet, snapshot);
                    _logger.Info("Kernel user data reflected to base home.");
                });
        }

        void IKernelUserDataReflectionService.ReflectToBaseHomeOnly(WorkbookContext context)
        {
            ReflectToBaseHomeOnly(context);
        }

        private void ReflectToBaseHome(string systemRoot, Excel.Workbook kernelWorkbook, Excel.Worksheet userDataWorksheet, UserDataSnapshot snapshot)
        {
            string baseWorkbookPath = WorkbookFileNameResolver.ResolveExistingBaseWorkbookPath(systemRoot, _pathCompatibilityService);
            if (!_pathCompatibilityService.FileExistsSafe(baseWorkbookPath))
            {
                throw new FileNotFoundException("Base workbook was not found.", baseWorkbookPath);
            }

            BaseHomeReflectionPlan reflectionPlan = BuildBaseHomeReflectionPlan(kernelWorkbook, userDataWorksheet, snapshot);
            Excel.Workbook baseWorkbook = _excelInteropService.FindOpenWorkbook(baseWorkbookPath);
            if (baseWorkbook != null)
            {
                ApplyBaseHomeReflectionPlan(baseWorkbook, reflectionPlan);
                baseWorkbook.Save();
                return;
            }

            ExecuteManagedHiddenReflectionSession(
                baseWorkbookPath,
                "Base",
                workbook => ApplyBaseHomeReflectionPlan(workbook, reflectionPlan));
        }

        private void ReflectToAccountingSet(Excel.Workbook kernelWorkbook, UserDataSnapshot snapshot)
        {
            string accountingWorkbookPath = _accountingTemplateResolver.ResolveTemplatePath(kernelWorkbook);
            AccountingReflectionPlan reflectionPlan = BuildAccountingReflectionPlan(
                snapshot,
                _pathCompatibilityService.NormalizePath(_excelInteropService.GetWorkbookFullName(kernelWorkbook)));
            Excel.Workbook accountingWorkbook = _excelInteropService.FindOpenWorkbook(accountingWorkbookPath);
            if (accountingWorkbook != null)
            {
                ApplyAccountingReflectionPlan(accountingWorkbook, reflectionPlan);
                accountingWorkbook.Save();
                return;
            }

            ExecuteManagedHiddenReflectionSession(
                accountingWorkbookPath,
                "Accounting",
                workbook => ApplyAccountingReflectionPlan(workbook, reflectionPlan));
        }

        private void ExecuteReflection(WorkbookContext context, Action<Excel.Workbook, string, Excel.Worksheet, UserDataSnapshot> action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            Excel.Workbook kernelWorkbook = ResolveReflectionKernelWorkbook(context);
            string systemRoot = ResolveReflectionSystemRoot(context, kernelWorkbook);

            Excel.Application application = kernelWorkbook.Application;
            ExcelApplicationUiState uiState = ExcelApplicationUiState.Capture(application);
            LogApplicationLifetimeOwnerFacts(
                "shared-current-app-boundary",
                ApplicationKindSharedCurrent,
                ApplicationLifetimeOwnerUserOrExcelHost,
                application,
                context == null ? string.Empty : context.SystemRoot);

            try
            {
                uiState.ApplyQuietMode(application);

                Excel.Worksheet userDataWorksheet = GetKernelUserDataWorksheet(kernelWorkbook);
                UserDataSnapshot snapshot = ReadUserDataSnapshot(userDataWorksheet, _excelInteropService);
                action(kernelWorkbook, systemRoot, userDataWorksheet, snapshot);
            }
            finally
            {
                uiState.Restore(application);
            }
        }

        private Excel.Workbook ResolveReflectionKernelWorkbook(WorkbookContext context)
        {
            if (context == null)
            {
                throw new InvalidOperationException("WorkbookContext is required for user-data reflection.");
            }

            Excel.Workbook kernelWorkbook = context.Workbook;
            if (kernelWorkbook == null)
            {
                throw new InvalidOperationException("Kernel workbook context was not available for user-data reflection.");
            }

            if (!_kernelWorkbookService.IsKernelWorkbook(kernelWorkbook))
            {
                throw new InvalidOperationException("Reflection target workbook must be a Kernel workbook.");
            }

            return kernelWorkbook;
        }
        private Excel.Worksheet GetKernelUserDataWorksheet(Excel.Workbook kernelWorkbook)
        {
            Excel.Worksheet worksheet = _excelInteropService.FindWorksheetByCodeName(kernelWorkbook, AccountingSetSpec.UserDataSheetCodeName);
            if (worksheet != null)
            {
                return worksheet;
            }

            try
            {
                worksheet = kernelWorkbook.Worksheets[AccountingSetSpec.UserDataSheetName] as Excel.Worksheet;
            }
            catch
            {
                worksheet = null;
            }

            if (worksheet == null)
            {
                throw new InvalidOperationException("Kernel user-data worksheet was not found.");
            }

            return worksheet;
        }

        private static Excel.Worksheet GetCaseHomeWorksheet(Excel.Workbook workbook)
        {
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                if (string.Equals(worksheet.CodeName, CaseHomeSheetCodeName, StringComparison.OrdinalIgnoreCase))
                {
                    return worksheet;
                }
            }

            try
            {
                Excel.Worksheet worksheet = workbook.Worksheets[CaseHomeSheetName] as Excel.Worksheet;
                if (worksheet != null)
                {
                    return worksheet;
                }
            }
            catch
            {
                // fallback to the first worksheet when named lookup fails.
            }

            Excel.Worksheet fallback = workbook.Worksheets[1] as Excel.Worksheet;
            if (fallback == null)
            {
                throw new InvalidOperationException("Case home worksheet was not found.");
            }

            return fallback;
        }

        private static UserDataSnapshot ReadUserDataSnapshot(Excel.Worksheet userDataWorksheet, ExcelInteropService excelInteropService)
        {
            if (userDataWorksheet == null)
            {
                throw new ArgumentNullException(nameof(userDataWorksheet));
            }

            if (excelInteropService == null)
            {
                throw new ArgumentNullException(nameof(excelInteropService));
            }

            IReadOnlyDictionary<string, string> values = excelInteropService.ReadKeyValueMapFromColumnsAandB(userDataWorksheet);
            if (values == null || values.Count == 0)
            {
                throw new InvalidOperationException("User-data values could not be read.");
            }

            return new UserDataSnapshot(values);
        }

        private string ResolveReflectionSystemRoot(WorkbookContext context, Excel.Workbook kernelWorkbook)
        {
            string expectedSystemRoot = _pathCompatibilityService.NormalizePath(context == null ? string.Empty : context.SystemRoot);
            if (string.IsNullOrWhiteSpace(expectedSystemRoot))
            {
                throw new InvalidOperationException("SYSTEM_ROOT context was not available for user-data reflection.");
            }

            string workbookSystemRoot = _pathCompatibilityService.NormalizePath(_excelInteropService.TryGetDocumentProperty(kernelWorkbook, "SYSTEM_ROOT"));
            if (string.IsNullOrWhiteSpace(workbookSystemRoot)
                || !string.Equals(workbookSystemRoot, expectedSystemRoot, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("Kernel workbook SYSTEM_ROOT mismatched for user-data reflection.");
            }

            return expectedSystemRoot;
        }

        private void ExecuteManagedHiddenReflectionSession(
            string workbookPath,
            string workbookKind,
            Action<Excel.Workbook> action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            if (string.IsNullOrWhiteSpace(workbookPath))
            {
                throw new InvalidOperationException(workbookKind + " workbook path could not be resolved.");
            }

            Excel.Application isolatedApplication = null;
            Excel.Workbook isolatedWorkbook = null;

            try
            {
                isolatedApplication = CreateHiddenIsolatedApplication();
                _logger.Info(
                    "Kernel user data reflection hidden session opening. target="
                    + workbookKind
                    + ", appHwnd="
                    + SafeApplicationHwnd(isolatedApplication)
                    + ", path="
                    + workbookPath
                    + ", "
                    + BuildApplicationOwnerFacts(ApplicationKindIsolated, ApplicationLifetimeOwnerKernelUserDataReflectionService));
                LogApplicationLifetimeOwnerFacts(
                    "isolated-app-boundary",
                    ApplicationKindIsolated,
                    ApplicationLifetimeOwnerKernelUserDataReflectionService,
                    isolatedApplication,
                    workbookPath);

                isolatedWorkbook = OpenWorkbookInManagedHiddenSession(isolatedApplication, workbookPath);
                if (isolatedWorkbook == null)
                {
                    throw new InvalidOperationException(workbookKind + " workbook could not be opened. path=" + workbookPath);
                }

                _accountingWorkbookService.SetWorkbookWindowsVisible(isolatedWorkbook, false);
                _logger.Info("Kernel user data reflection hidden workbook opened in managed session. target=" + workbookKind + ", path=" + workbookPath);

                action(isolatedWorkbook);
                RestoreOwnedWorkbookWindowVisibilityForSave(isolatedWorkbook, workbookKind, workbookPath);
                isolatedWorkbook.Save();
            }
            finally
            {
                bool workbookPresent = isolatedWorkbook != null;
                bool workbookCloseCompleted = !workbookPresent;
                bool applicationPresent = isolatedApplication != null;
                bool applicationQuitAttempted = false;
                bool applicationQuitCompleted = !applicationPresent;

                if (isolatedWorkbook != null)
                {
                    _logger.Info(
                        "Kernel user data reflection hidden workbook closing in managed session. target="
                        + workbookKind
                        + ", appHwnd="
                        + SafeApplicationHwnd(isolatedApplication)
                        + ", path="
                        + workbookPath);
                    workbookCloseCompleted = CloseWorkbookQuietly(isolatedWorkbook, workbookKind, workbookPath);
                }

                if (isolatedApplication != null)
                {
                    string applicationHwnd = SafeApplicationHwnd(isolatedApplication);
                    applicationQuitAttempted = true;
                    applicationQuitCompleted = QuitApplicationQuietly(isolatedApplication, workbookKind, workbookPath);
                    _logger.Info("Kernel user data reflection hidden session quit. target=" + workbookKind + ", appHwnd=" + applicationHwnd + ", path=" + workbookPath);
                }

                LogManagedHiddenReflectionCleanupOutcome(
                    workbookKind,
                    workbookPath,
                    workbookPresent,
                    workbookCloseCompleted,
                    applicationPresent,
                    applicationQuitAttempted,
                    applicationQuitCompleted);
            }
        }

        private static Excel.Application CreateHiddenIsolatedApplication()
        {
            return new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false,
                ScreenUpdating = false,
                EnableEvents = false
            };
        }

        private static Excel.Workbook OpenWorkbookInManagedHiddenSession(Excel.Application application, string workbookPath)
        {
            return application.Workbooks.Open(workbookPath, UpdateLinks: 0, ReadOnly: false);
        }

        private void RestoreOwnedWorkbookWindowVisibilityForSave(Excel.Workbook workbook, string workbookKind, string workbookPath)
        {
            if (workbook == null)
            {
                return;
            }

            _accountingWorkbookService.SetWorkbookWindowsVisible(workbook, true);
            _logger.Info(
                "Kernel user data reflection owned workbook windows restored before save. target="
                + workbookKind
                + ", path="
                + workbookPath);
        }

        private bool CloseWorkbookQuietly(Excel.Workbook workbook, string workbookKind, string workbookPath)
        {
            if (workbook == null)
            {
                return false;
            }

            bool completed = false;
            try
            {
                WorkbookCloseInteropHelper.CloseOwnedWorkbookWithoutSave(
                    workbook,
                    _logger,
                    nameof(KernelUserDataReflectionService) + ".CloseWorkbookQuietly target=" + (workbookKind ?? string.Empty));
                completed = true;
            }
            catch
            {
                // owned workbook cleanup failure must not mask the original exception.
            }
            finally
            {
                CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.FinalRelease(
                    workbook,
                    _logger,
                    nameof(KernelUserDataReflectionService)
                    + ".CloseWorkbookQuietly target="
                    + (workbookKind ?? string.Empty)
                    + ", path="
                    + (workbookPath ?? string.Empty));
            }

            return completed;
        }

        private bool QuitApplicationQuietly(Excel.Application application, string workbookKind, string workbookPath)
        {
            if (application == null)
            {
                return false;
            }

            bool completed = false;
            try
            {
                application.Quit();
                completed = true;
            }
            catch
            {
                // isolated application cleanup failure must not mask the original exception.
            }
            finally
            {
                CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.FinalRelease(
                    application,
                    _logger,
                    nameof(KernelUserDataReflectionService)
                    + ".QuitApplicationQuietly target="
                    + (workbookKind ?? string.Empty)
                    + ", appHwnd="
                    + SafeApplicationHwnd(application)
                    + ", path="
                    + (workbookPath ?? string.Empty));
            }

            return completed;
        }

        private void LogManagedHiddenReflectionCleanupOutcome(
            string workbookKind,
            string workbookPath,
            bool workbookPresent,
            bool workbookCloseCompleted,
            bool applicationPresent,
            bool applicationQuitAttempted,
            bool applicationQuitCompleted)
        {
            string hiddenOutcome = ResolveReflectionHiddenCleanupOutcome(
                workbookPresent,
                workbookCloseCompleted,
                applicationPresent,
                applicationQuitCompleted);
            string isolatedOutcome = ResolveReflectionIsolatedAppOutcome(
                applicationPresent,
                applicationQuitAttempted,
                applicationQuitCompleted);
            _logger.Info(
                "[KernelFlickerTrace] source=KernelUserDataReflectionService action=hidden-excel-cleanup-outcome"
                + " target=" + (workbookKind ?? string.Empty)
                + ", path=" + (workbookPath ?? string.Empty)
                + ", " + BuildApplicationOwnerFacts(ApplicationKindIsolated, ApplicationLifetimeOwnerKernelUserDataReflectionService)
                + ", hiddenCleanupOutcome=" + hiddenOutcome
                + ", isolatedAppOutcome=" + isolatedOutcome
                + ", workbookPresent=" + workbookPresent.ToString()
                + ", workbookCloseCompleted=" + workbookCloseCompleted.ToString()
                + ", appPresent=" + applicationPresent.ToString()
                + ", appQuitAttempted=" + applicationQuitAttempted.ToString()
                + ", appQuitCompleted=" + applicationQuitCompleted.ToString()
                + ", owner=KernelUserDataReflectionService");
        }

        private void LogApplicationLifetimeOwnerFacts(
            string action,
            string applicationKind,
            string applicationLifetimeOwner,
            Excel.Application application,
            string pathOrSystemRoot)
        {
            _logger.Info(
                "[KernelFlickerTrace] source=KernelUserDataReflectionService"
                + " action=" + (action ?? string.Empty)
                + " " + BuildApplicationOwnerFacts(applicationKind, applicationLifetimeOwner)
                + ", appHwnd=" + SafeApplicationHwnd(application)
                + ", pathOrSystemRoot=" + (pathOrSystemRoot ?? string.Empty));
        }

        private static string BuildApplicationOwnerFacts(string applicationKind, string applicationLifetimeOwner)
        {
            bool isSharedCurrentApp = string.Equals(applicationKind, ApplicationKindSharedCurrent, StringComparison.Ordinal);
            bool isIsolatedApp = string.Equals(applicationKind, ApplicationKindIsolated, StringComparison.Ordinal);
            bool isRetainedHiddenAppCache = string.Equals(applicationKind, ApplicationKindRetainedHiddenAppCache, StringComparison.Ordinal);
            return "applicationKind=" + (applicationKind ?? string.Empty)
                + ",applicationLifetimeOwner=" + (applicationLifetimeOwner ?? string.Empty)
                + ",isSharedCurrentApp=" + isSharedCurrentApp.ToString()
                + ",isIsolatedApp=" + isIsolatedApp.ToString()
                + ",isRetainedHiddenAppCache=" + isRetainedHiddenAppCache.ToString();
        }

        private static string ResolveReflectionHiddenCleanupOutcome(
            bool workbookPresent,
            bool workbookCloseCompleted,
            bool applicationPresent,
            bool applicationQuitCompleted)
        {
            if (!workbookPresent && !applicationPresent)
            {
                return HiddenExcelCleanupNotRequired;
            }

            return (!workbookPresent || workbookCloseCompleted) && (!applicationPresent || applicationQuitCompleted)
                ? HiddenExcelCleanupCompleted
                : HiddenExcelCleanupDegraded;
        }

        private static string ResolveReflectionIsolatedAppOutcome(bool applicationPresent, bool applicationQuitAttempted, bool applicationQuitCompleted)
        {
            if (!applicationPresent)
            {
                return IsolatedAppReleaseNotRequired;
            }

            if (applicationQuitCompleted)
            {
                return IsolatedAppReleased;
            }

            return applicationQuitAttempted
                ? IsolatedAppReleaseFailed
                : IsolatedAppReleaseDegraded;
        }

        private static string SafeApplicationHwnd(Excel.Application application)
        {
            try
            {
                return application == null ? string.Empty : Convert.ToString(application.Hwnd) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void ApplyBaseHomeReflectionPlan(Excel.Workbook baseWorkbook, BaseHomeReflectionPlan reflectionPlan)
        {
            if (baseWorkbook == null)
            {
                throw new ArgumentNullException(nameof(baseWorkbook));
            }

            if (reflectionPlan == null)
            {
                throw new ArgumentNullException(nameof(reflectionPlan));
            }

            Excel.Worksheet homeWorksheet = GetCaseHomeWorksheet(baseWorkbook);
            bool wasProtected = homeWorksheet.ProtectContents || homeWorksheet.ProtectDrawingObjects || homeWorksheet.ProtectScenarios;
            if (wasProtected)
            {
                homeWorksheet.Unprotect(UnprotectPassword);
            }

            ApplyBaseHomeKeyValues(homeWorksheet, reflectionPlan);

            if (wasProtected)
            {
                Excel.Range allCells = homeWorksheet.Cells as Excel.Range;
                if (allCells != null)
                {
                    allCells.Locked = true;
                }

                Excel.Range unlockedColumn = homeWorksheet.Columns["B"] as Excel.Range;
                if (unlockedColumn != null)
                {
                    unlockedColumn.Locked = false;
                }

                homeWorksheet.Protect(
                    Password: UnprotectPassword,
                    UserInterfaceOnly: true,
                    AllowFiltering: true,
                    AllowSorting: true);
                homeWorksheet.EnableSelection = Excel.XlEnableSelection.xlUnlockedCells;
            }
        }

        private static AccountingReflectionPlan BuildAccountingReflectionPlan(UserDataSnapshot snapshot, string sourceKernelPath)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException(nameof(snapshot));
            }

            return new AccountingReflectionPlan(
                sourceKernelPath,
                BuildUserAddressLine(snapshot, false),
                BuildUserAddressLine(snapshot, true),
                snapshot.GetValue(UserDataAccountingNameRow1Key),
                snapshot.GetValue(UserDataAccountingNameRow2Key));
        }

        private void ApplyAccountingReflectionPlan(Excel.Workbook accountingWorkbook, AccountingReflectionPlan reflectionPlan)
        {
            if (accountingWorkbook == null)
            {
                throw new ArgumentNullException(nameof(accountingWorkbook));
            }

            if (reflectionPlan == null)
            {
                throw new ArgumentNullException(nameof(reflectionPlan));
            }

            _excelInteropService.SetDocumentProperty(
                accountingWorkbook,
                AccountingSetSpec.SourceKernelPathPropertyName,
                reflectionPlan.SourceKernelPath);

            _accountingWorkbookService.WriteCell(accountingWorkbook, AccountingSetSpec.InvoiceSheetName, AccountingSetSpec.InvoiceNameRow1CellAddress, reflectionPlan.NameRow1);
            _accountingWorkbookService.WriteCell(accountingWorkbook, AccountingSetSpec.InvoiceSheetName, AccountingSetSpec.InvoiceNameRow2CellAddress, reflectionPlan.NameRow2);
            _accountingWorkbookService.WriteCell(accountingWorkbook, AccountingSetSpec.InstallmentSheetName, AccountingSetSpec.InstallmentNameRow1CellAddress, reflectionPlan.NameRow1);
            _accountingWorkbookService.WriteCell(accountingWorkbook, AccountingSetSpec.InstallmentSheetName, AccountingSetSpec.InstallmentNameRow2CellAddress, reflectionPlan.NameRow2);
            _accountingWorkbookService.WriteCell(accountingWorkbook, AccountingSetSpec.InstallmentSheetName, AccountingSetSpec.InstallmentAddressCellAddress, reflectionPlan.AddressLineWithBreak);
            _accountingWorkbookService.WriteSameValueToSheets(
                accountingWorkbook,
                new[]
                {
                    AccountingSetSpec.EstimateSheetName,
                    AccountingSetSpec.InvoiceSheetName,
                    AccountingSetSpec.ReceiptSheetName
                },
                AccountingSetSpec.AccountingAddressCellAddress,
                reflectionPlan.AddressLine);
            _accountingWorkbookService.WriteSameValueToSheets(
                accountingWorkbook,
                new[]
                {
                    AccountingSetSpec.InstallmentSheetName,
                    AccountingSetSpec.PaymentHistorySheetName
                },
                AccountingSetSpec.InstallmentAddressCellAddress,
                reflectionPlan.AddressLineWithBreak);
        }

        private static string BuildUserAddressLine(UserDataSnapshot snapshot, bool withLineBreak)
        {
            string b1 = snapshot.GetValue(UserDataPostalCodeKey).Trim();
            string b2 = snapshot.GetValue(UserDataAddressKey).Trim();
            string b3 = snapshot.GetValue(UserDataOfficeNameKey).Trim();
            string b4 = snapshot.GetValue(UserDataPhoneKey).Trim();
            string fullWidthSpace = "\u3000";

            if (withLineBreak)
            {
                return "\u3012" + b1 + fullWidthSpace + b2 + Environment.NewLine
                    + new string('\u3000', 11) + b3 + fullWidthSpace + "\u2121" + b4;
            }

            return "\u3012" + b1 + fullWidthSpace + b2 + fullWidthSpace + b3 + fullWidthSpace + "\u2121" + b4;
        }

        /// <summary>
        /// UserData シートの値を保持するスナップショット。
        /// </summary>
        private sealed class UserDataSnapshot
        {
            private readonly IReadOnlyDictionary<string, string> _values;

            internal UserDataSnapshot(IReadOnlyDictionary<string, string> values)
            {
                _values = values ?? throw new ArgumentNullException(nameof(values));
                if (_values.Count == 0)
                {
                    throw new ArgumentException("User-data values were empty.", nameof(values));
                }
            }

            internal string GetValue(string key)
            {
                return FieldKeyRenameMap.GetValueWithAliases(_values, key);
            }
        }

        private sealed class BaseHomeReflectionPlan
        {
            internal BaseHomeReflectionPlan(IReadOnlyDictionary<string, string> homeValuesByKey)
            {
                var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (homeValuesByKey == null)
                {
                    HomeValuesByKey = values;
                    return;
                }

                foreach (KeyValuePair<string, string> pair in homeValuesByKey)
                {
                    values[pair.Key] = pair.Value ?? string.Empty;
                }

                HomeValuesByKey = values;
            }

            internal IReadOnlyDictionary<string, string> HomeValuesByKey { get; }
        }

        private BaseHomeReflectionPlan BuildBaseHomeReflectionPlan(
            Excel.Workbook kernelWorkbook,
            Excel.Worksheet userDataWorksheet,
            UserDataSnapshot snapshot)
        {
            if (kernelWorkbook == null)
            {
                throw new ArgumentNullException(nameof(kernelWorkbook));
            }

            if (userDataWorksheet == null)
            {
                throw new ArgumentNullException(nameof(userDataWorksheet));
            }

            if (snapshot == null)
            {
                throw new ArgumentNullException(nameof(snapshot));
            }

            IReadOnlyList<UserDataBaseMappingDefinition> mappingDefinitions = _userDataBaseMappingRepository.LoadEnabledDefinitions(kernelWorkbook);
            if (mappingDefinitions == null || mappingDefinitions.Count == 0)
            {
                throw new InvalidOperationException("Kernel の管理シート UserData_BaseMapping を読み取れません。");
            }

            IReadOnlyDictionary<string, string> userDataValues = _excelInteropService.ReadKeyValueMapFromColumnsAandB(userDataWorksheet);
            var homeValuesByKey = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (UserDataBaseMappingDefinition mapping in mappingDefinitions)
            {
                if (mapping == null)
                {
                    continue;
                }

                string sourceFieldKey = FieldKeyRenameMap.NormalizeToCurrent(mapping.SourceFieldKey);
                string targetFieldKey = FieldKeyRenameMap.NormalizeToCurrent(mapping.TargetFieldKey);
                string sourceValue;
                if (!FieldKeyRenameMap.TryGetValueWithAliases(userDataValues, sourceFieldKey, out sourceValue))
                {
                    throw new InvalidOperationException("Kernel ユーザー情報にキーが見つかりません。 key=" + mapping.SourceFieldKey);
                }

                homeValuesByKey[targetFieldKey] = sourceValue ?? string.Empty;
            }

            return new BaseHomeReflectionPlan(homeValuesByKey);
        }

        private void ApplyBaseHomeKeyValues(Excel.Worksheet homeWorksheet, BaseHomeReflectionPlan reflectionPlan)
        {
            if (homeWorksheet == null)
            {
                throw new ArgumentNullException(nameof(homeWorksheet));
            }

            if (reflectionPlan == null)
            {
                throw new ArgumentNullException(nameof(reflectionPlan));
            }

            Excel.Range keyRange = null;
            Excel.Range targetCell = null;
            try
            {
                Excel.Range lastCell = homeWorksheet.Cells[homeWorksheet.Rows.Count, "A"] as Excel.Range;
                if (lastCell == null)
                {
                    throw new InvalidOperationException("Base HOME の最終行セルを取得できませんでした。");
                }

                Excel.Range lastUsedCell = lastCell.End[Excel.XlDirection.xlUp] as Excel.Range;
                if (lastUsedCell == null)
                {
                    throw new InvalidOperationException("Base HOME の使用済み最終行を取得できませんでした。");
                }

                int lastRow = lastUsedCell.Row;
                if (lastRow < 1)
                {
                    throw new InvalidOperationException("Base HOME のキー列を読み取れませんでした。");
                }

                keyRange = homeWorksheet.Range["A1", "A" + lastRow.ToString()];
                object[,] keyValues = keyRange.Value2 as object[,];
                if (keyValues == null)
                {
                    throw new InvalidOperationException("Base HOME のキー一覧を取得できませんでした。");
                }

                var rowByKey = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                int upperRow = keyValues.GetUpperBound(0);
                for (int rowIndex = 1; rowIndex <= upperRow; rowIndex++)
                {
                    string key = (Convert.ToString(keyValues[rowIndex, 1]) ?? string.Empty).Trim();
                    if (key.Length == 0 || rowByKey.ContainsKey(key))
                    {
                        continue;
                    }

                    rowByKey[key] = rowIndex;
                }

                foreach (KeyValuePair<string, string> pair in reflectionPlan.HomeValuesByKey)
                {
                    int rowNumber;
                    if (!rowByKey.TryGetValue(pair.Key, out rowNumber))
                    {
                        throw new InvalidOperationException("Base HOME にキーが見つかりません。 key=" + pair.Key);
                    }

                    targetCell = homeWorksheet.Cells[rowNumber, "B"] as Excel.Range;
                    if (targetCell == null)
                    {
                        throw new InvalidOperationException("Base HOME の転記先セルを取得できません。 key=" + pair.Key);
                    }

                    _accountingWorkbookService.WriteCellValue(homeWorksheet, "B" + rowNumber.ToString(), pair.Value ?? string.Empty);
                    CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.Release(targetCell, _logger);
                    targetCell = null;
                }
            }
            finally
            {
                CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.Release(targetCell, _logger);
                CaseInfoSystem.ExcelAddIn.Infrastructure.ComObjectReleaseService.Release(keyRange, _logger);
            }
        }

        private sealed class AccountingReflectionPlan
        {
            internal AccountingReflectionPlan(
                string sourceKernelPath,
                string addressLine,
                string addressLineWithBreak,
                string nameRow1,
                string nameRow2)
            {
                SourceKernelPath = sourceKernelPath ?? string.Empty;
                AddressLine = addressLine ?? string.Empty;
                AddressLineWithBreak = addressLineWithBreak ?? string.Empty;
                NameRow1 = nameRow1 ?? string.Empty;
                NameRow2 = nameRow2 ?? string.Empty;
            }

            internal string SourceKernelPath { get; }

            internal string AddressLine { get; }

            internal string AddressLineWithBreak { get; }

            internal string NameRow1 { get; }

            internal string NameRow2 { get; }
        }

        /// <summary>
        /// Excel UI の状態を退避・復元するためのスナップショット。
        /// </summary>
        private sealed class ExcelApplicationUiState
        {
            private ExcelApplicationUiState()
            {
            }

            internal bool ScreenUpdating { get; private set; }

            internal bool EnableEvents { get; private set; }

            internal bool DisplayAlerts { get; private set; }

            internal object StatusBar { get; private set; }

            internal static ExcelApplicationUiState Capture(Excel.Application application)
            {
                if (application == null)
                {
                    return new ExcelApplicationUiState();
                }

                return new ExcelApplicationUiState
                {
                    ScreenUpdating = application.ScreenUpdating,
                    EnableEvents = application.EnableEvents,
                    DisplayAlerts = application.DisplayAlerts,
                    StatusBar = application.StatusBar
                };
            }

            internal void ApplyQuietMode(Excel.Application application)
            {
                if (application == null)
                {
                    return;
                }

                try
                {
                    application.ScreenUpdating = false;
                    application.EnableEvents = false;
                    application.DisplayAlerts = false;
                    application.StatusBar = "ユーザー情報を反映中...";
                }
                catch
                {
                    // UI quiet-mode failures must not stop reflection.
                }
            }

            internal void Restore(Excel.Application application)
            {
                if (application == null)
                {
                    return;
                }

                try
                {
                    application.ScreenUpdating = ScreenUpdating;
                    application.EnableEvents = EnableEvents;
                    application.DisplayAlerts = DisplayAlerts;
                    application.StatusBar = StatusBar;
                }
                catch
                {
                    // UI restore failures must not stop shutdown paths.
                }
            }
        }

    }
}



