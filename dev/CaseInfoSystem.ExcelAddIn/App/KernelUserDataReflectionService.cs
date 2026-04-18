using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    /// Kernel のユーザー情報を Base HOME と会計セットへ反映するサービス。
    /// shUserData の値を読み取り、転記先ブックへ静かに書き戻す。
    /// </summary>
    internal sealed class KernelUserDataReflectionService
    {
        private const string CaseHomeSheetCodeName = "shHOME";
        private const string CaseHomeSheetName = "ホーム";
        private const string UnprotectPassword = "";
        private const string UserDataPostalCodeKey = "当方_郵便番号";
        private const string UserDataAddressKey = "当方_住所";
        private const string UserDataOfficeNameKey = "当方_事務所名";
        private const string UserDataPhoneKey = "当方_電話";
        private const string UserDataAccountingNameRow1Key = "銀行・支店";
        private const string UserDataAccountingNameRow2Key = "口座番号・名義";

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
        internal void ReflectAll()
        {
            ExecuteReflection(
                (kernelWorkbook, userDataWorksheet, snapshot) =>
                {
                    ReflectToAccountingSet(kernelWorkbook, snapshot);
                    ReflectToBaseHome(kernelWorkbook, userDataWorksheet, snapshot);
                    _logger.Info("Kernel user data reflected to accounting set and base home.");
                });
        }

        internal void ReflectToAccountingSetOnly()
        {
            ExecuteReflection(
                (kernelWorkbook, userDataWorksheet, snapshot) =>
                {
                    ReflectToAccountingSet(kernelWorkbook, snapshot);
                    _logger.Info("Kernel user data reflected to accounting set.");
                });
        }

        internal void ReflectToBaseHomeOnly()
        {
            ExecuteReflection(
                (kernelWorkbook, userDataWorksheet, snapshot) =>
                {
                    ReflectToBaseHome(kernelWorkbook, userDataWorksheet, snapshot);
                    _logger.Info("Kernel user data reflected to base home.");
                });
        }

        private void ReflectToBaseHome(Excel.Workbook kernelWorkbook, Excel.Worksheet userDataWorksheet, UserDataSnapshot snapshot)
        {
            string systemRoot = ResolveSystemRoot(kernelWorkbook);
            string baseWorkbookPath = WorkbookFileNameResolver.ResolveExistingBaseWorkbookPath(systemRoot, _pathCompatibilityService);
            if (!_pathCompatibilityService.FileExistsSafe(baseWorkbookPath))
            {
                throw new FileNotFoundException("Base workbook was not found.", baseWorkbookPath);
            }

            Excel.Workbook baseWorkbook = _excelInteropService.FindOpenWorkbook(baseWorkbookPath);
            bool wasAlreadyOpen = baseWorkbook != null;
            HiddenWorkbookSession hiddenSession = null;

            try
            {
                if (baseWorkbook == null)
                {
                    hiddenSession = OpenWorkbookHiddenInIsolatedApplication(baseWorkbookPath);
                    baseWorkbook = hiddenSession.Workbook;
                    _logger.Info("Kernel user data reflection hidden workbook opened for Base in isolated application. path=" + baseWorkbookPath);
                }

                BaseHomeReflectionPlan reflectionPlan = BuildBaseHomeReflectionPlan(kernelWorkbook, userDataWorksheet, snapshot);
                ApplyBaseHomeReflectionPlan(baseWorkbook, reflectionPlan);
                baseWorkbook.Save();
            }
            finally
            {
                if (!wasAlreadyOpen && hiddenSession != null)
                {
                    hiddenSession.CloseWithoutSaving();
                    _logger.Info("Kernel user data reflection hidden workbook closed for Base in isolated application. path=" + baseWorkbookPath);
                }
            }
        }

        private void ReflectToAccountingSet(Excel.Workbook kernelWorkbook, UserDataSnapshot snapshot)
        {
            string accountingWorkbookPath = _accountingTemplateResolver.ResolveTemplatePath(kernelWorkbook);
            Excel.Workbook accountingWorkbook = _excelInteropService.FindOpenWorkbook(accountingWorkbookPath);
            bool wasAlreadyOpen = accountingWorkbook != null;
            HiddenWorkbookSession hiddenSession = null;

            try
            {
                if (accountingWorkbook == null)
                {
                    hiddenSession = OpenWorkbookHiddenInIsolatedApplication(accountingWorkbookPath);
                    accountingWorkbook = hiddenSession.Workbook;
                    _logger.Info("Kernel user data reflection hidden workbook opened for Accounting in isolated application. path=" + accountingWorkbookPath);
                }

                AccountingReflectionPlan reflectionPlan = BuildAccountingReflectionPlan(
                    snapshot,
                    _pathCompatibilityService.NormalizePath(_excelInteropService.GetWorkbookFullName(kernelWorkbook)));
                ApplyAccountingReflectionPlan(accountingWorkbook, reflectionPlan);
                accountingWorkbook.Save();
            }
            finally
            {
                if (!wasAlreadyOpen && hiddenSession != null)
                {
                    hiddenSession.CloseWithoutSaving();
                    _logger.Info("Kernel user data reflection hidden workbook closed for Accounting in isolated application. path=" + accountingWorkbookPath);
                }
            }
        }

        private void ExecuteReflection(Action<Excel.Workbook, Excel.Worksheet, UserDataSnapshot> action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            Excel.Workbook kernelWorkbook = _kernelWorkbookService.GetOpenKernelWorkbook();
            if (kernelWorkbook == null)
            {
                throw new InvalidOperationException("Kernel workbook is not open.");
            }

            Excel.Application application = kernelWorkbook.Application;
            ExcelApplicationUiState uiState = ExcelApplicationUiState.Capture(application);

            try
            {
                uiState.ApplyQuietMode(application);

                Excel.Worksheet userDataWorksheet = GetKernelUserDataWorksheet(kernelWorkbook);
                UserDataSnapshot snapshot = ReadUserDataSnapshot(userDataWorksheet, _excelInteropService);
                action(kernelWorkbook, userDataWorksheet, snapshot);
            }
            finally
            {
                uiState.Restore(application);
            }
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

        private string ResolveSystemRoot(Excel.Workbook kernelWorkbook)
        {
            string systemRoot = _pathCompatibilityService.NormalizePath(_excelInteropService.TryGetDocumentProperty(kernelWorkbook, "SYSTEM_ROOT"));
            if (!string.IsNullOrWhiteSpace(systemRoot))
            {
                return systemRoot;
            }

            systemRoot = _pathCompatibilityService.NormalizePath(_excelInteropService.GetWorkbookPath(kernelWorkbook));
            if (string.IsNullOrWhiteSpace(systemRoot))
            {
                throw new InvalidOperationException("SYSTEM_ROOT could not be resolved.");
            }

            return systemRoot;
        }

        private HiddenWorkbookSession OpenWorkbookHiddenInIsolatedApplication(string workbookPath)
        {
            if (string.IsNullOrWhiteSpace(workbookPath))
            {
                throw new ArgumentException("Workbook path is required.", nameof(workbookPath));
            }

            Excel.Application hiddenApplication = new Excel.Application
            {
                Visible = false,
                DisplayAlerts = false,
                ScreenUpdating = false,
                EnableEvents = false
            };

            Excel.Workbook workbook = null;
            try
            {
                workbook = hiddenApplication.Workbooks.Open(workbookPath, UpdateLinks: 0, ReadOnly: false);
                _accountingWorkbookService.SetWorkbookWindowsVisible(workbook, false);
                return new HiddenWorkbookSession(hiddenApplication, workbook);
            }
            catch
            {
                if (workbook != null)
                {
                    try
                    {
                        workbook.Close(SaveChanges: false);
                    }
                    catch
                    {
                        // hidden workbook cleanup failure must not mask the original exception.
                    }
                }

                try
                {
                    hiddenApplication.Quit();
                }
                catch
                {
                    // hidden application cleanup failure must not mask the original exception.
                }

                ReleaseComObject(workbook);
                ReleaseComObject(hiddenApplication);
                throw;
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
                homeWorksheet.Cells.Locked = true;
                homeWorksheet.Columns["B"].Locked = false;
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
                if (string.IsNullOrWhiteSpace(key))
                {
                    return string.Empty;
                }

                return _values.TryGetValue(key, out string value) ? (value ?? string.Empty) : string.Empty;
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

                string sourceValue;
                if (!userDataValues.TryGetValue(mapping.SourceFieldKey, out sourceValue))
                {
                    throw new InvalidOperationException("Kernel ユーザー情報にキーが見つかりません。 key=" + mapping.SourceFieldKey);
                }

                homeValuesByKey[mapping.TargetFieldKey] = sourceValue ?? string.Empty;
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
                int lastRow = homeWorksheet.Cells[homeWorksheet.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;
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
                    ReleaseComObject(targetCell);
                    targetCell = null;
                }
            }
            finally
            {
                ReleaseComObject(targetCell);
                ReleaseComObject(keyRange);
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

        private sealed class HiddenWorkbookSession
        {
            internal HiddenWorkbookSession(Excel.Application application, Excel.Workbook workbook)
            {
                Application = application ?? throw new ArgumentNullException(nameof(application));
                Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            }

            internal Excel.Application Application { get; }

            internal Excel.Workbook Workbook { get; }

            internal void CloseWithoutSaving()
            {
                try
                {
                    Workbook.Close(SaveChanges: false);
                }
                finally
                {
                    try
                    {
                        Application.Quit();
                    }
                    finally
                    {
                        ReleaseComObject(Workbook);
                        ReleaseComObject(Application);
                    }
                }
            }
        }

        private static void ReleaseComObject(object comObject)
        {
            if (comObject == null || !Marshal.IsComObject(comObject))
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch
            {
                // COM 解放失敗は後続処理を止めない。
            }
        }
    }
}



