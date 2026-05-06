using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelWorkbookBindingService
    {
        private readonly Excel.Application _application;
        private readonly ExcelInteropService _excelInteropService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly Logger _logger;
        private readonly KernelOpenWorkbookLocator _kernelOpenWorkbookLocator;
        private readonly KernelWorkbookStateService _kernelWorkbookStateService;
        private readonly KernelWorkbookSettingsService _kernelWorkbookSettingsService;
        private KernelHomeBinding _homeBinding;

        internal KernelWorkbookBindingService(
            Excel.Application application,
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            Logger logger,
            KernelWorkbookService.KernelWorkbookServiceTestHooks testHooks = null)
        {
            _application = application;
            _excelInteropService = excelInteropService;
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _kernelOpenWorkbookLocator = new KernelOpenWorkbookLocator(
                _application,
                _excelInteropService,
                _pathCompatibilityService,
                _logger,
                resolveKernelWorkbookPathOverride: testHooks == null ? null : testHooks.ResolveKernelWorkbookPath,
                findOpenWorkbookOverride: testHooks == null ? null : testHooks.FindOpenWorkbook);
            _kernelWorkbookStateService = new KernelWorkbookStateService(
                _application,
                _excelInteropService,
                _logger,
                _kernelOpenWorkbookLocator,
                hasOtherVisibleWorkbookOverride: testHooks == null ? null : testHooks.HasOtherVisibleWorkbook,
                hasOtherWorkbookOverride: testHooks == null ? null : testHooks.HasOtherWorkbook);
            _kernelWorkbookSettingsService = new KernelWorkbookSettingsService();
        }

        internal bool IsKernelWorkbook(Excel.Workbook workbook)
        {
            return _kernelWorkbookStateService.IsKernelWorkbook(workbook);
        }

        internal Excel.Workbook ResolveKernelWorkbook(WorkbookContext context)
        {
            return _kernelOpenWorkbookLocator.ResolveKernelWorkbook(context);
        }

        internal Excel.Workbook ResolveKernelWorkbook(string systemRoot)
        {
            return _kernelOpenWorkbookLocator.ResolveKernelWorkbook(systemRoot);
        }

        internal bool BindHomeWorkbook(WorkbookContext context)
        {
            KernelHomeBinding binding = CreateHomeBinding(context);
            if (binding == null)
            {
                ClearHomeWorkbookBinding("BindHomeWorkbook.Invalid");
                _logger.Warn(
                    "Kernel HOME binding was not created from context. workbook="
                    + GetWorkbookFullName(context == null ? null : context.Workbook)
                    + ", contextSystemRoot="
                    + (context == null ? string.Empty : context.SystemRoot ?? string.Empty));
                return false;
            }

            _homeBinding = binding;
            _logger.Info(
                "Kernel HOME binding created. workbook="
                + GetWorkbookFullName(binding.Workbook)
                + ", systemRoot="
                + binding.SystemRoot);
            return true;
        }

        internal bool BindHomeWorkbook(Excel.Workbook workbook)
        {
            KernelHomeBinding binding = CreateHomeBinding(workbook);
            if (binding == null)
            {
                ClearHomeWorkbookBinding("BindHomeWorkbook.WorkbookInvalid");
                _logger.Warn(
                    "Kernel HOME binding was not created from workbook. workbook="
                    + GetWorkbookFullName(workbook));
                return false;
            }

            _homeBinding = binding;
            _logger.Info(
                "Kernel HOME binding created from workbook. workbook="
                + GetWorkbookFullName(binding.Workbook)
                + ", systemRoot="
                + binding.SystemRoot);
            return true;
        }

        internal void ClearHomeWorkbookBinding(string reason)
        {
            if (_homeBinding == null)
            {
                return;
            }

            _logger.Info(
                "Kernel HOME binding cleared. reason="
                + (reason ?? string.Empty)
                + ", workbook="
                + GetWorkbookFullName(_homeBinding.Workbook)
                + ", systemRoot="
                + (_homeBinding.SystemRoot ?? string.Empty));
            _homeBinding = null;
        }

        internal bool HasValidHomeWorkbookBinding()
        {
            Excel.Workbook workbook;
            return ResolveHomeBindingStatus("HasValidHomeWorkbookBinding", out workbook) == KernelHomeBindingStatus.Valid;
        }

        internal bool TryGetValidHomeWorkbookBinding(out Excel.Workbook workbook, out string systemRoot)
        {
            workbook = null;
            systemRoot = string.Empty;

            KernelHomeBindingStatus bindingStatus = ResolveHomeBindingStatus("TryGetValidHomeWorkbookBinding", out workbook);
            if (bindingStatus != KernelHomeBindingStatus.Valid)
            {
                workbook = null;
                return false;
            }

            systemRoot = _homeBinding == null ? string.Empty : _homeBinding.SystemRoot ?? string.Empty;
            if (string.IsNullOrWhiteSpace(systemRoot))
            {
                workbook = null;
                return false;
            }

            return true;
        }

        internal bool ShouldShowHomeOnStartup(Excel.Workbook startupWorkbook = null)
        {
            return _kernelWorkbookStateService.ShouldShowHomeOnStartup(startupWorkbook);
        }

        internal string DescribeStartupState()
        {
            return _kernelWorkbookStateService.DescribeStartupState();
        }

        internal KernelSettingsState LoadSettings()
        {
            Excel.Workbook workbook;
            KernelHomeBindingStatus bindingStatus = ResolveHomeBindingStatus("LoadSettings", out workbook);
            if (bindingStatus == KernelHomeBindingStatus.Invalid)
            {
                _logger.Warn("Kernel settings load failed closed because HOME binding was invalid.");
                return new KernelSettingsState();
            }

            if (bindingStatus == KernelHomeBindingStatus.None)
            {
                _logger.Warn("Kernel settings load failed closed because HOME binding was not available.");
                return new KernelSettingsState();
            }

            if (workbook == null)
            {
                _logger.Warn("Kernel settings load skipped because kernel workbook was not available.");
                return new KernelSettingsState();
            }

            string nameRuleA = _kernelWorkbookSettingsService.LoadNameRuleA(workbook);
            string nameRuleB = _kernelWorkbookSettingsService.LoadNameRuleB(workbook);
            string defaultRoot = _kernelWorkbookSettingsService.LoadDefaultRoot(workbook);
            string systemRoot = _excelInteropService == null
                ? string.Empty
                : KernelWorkbookResolver.TryGetKernelWorkbookDirectory(workbook, _excelInteropService, _pathCompatibilityService, _logger, IsKernelWorkbook);

            _logger.Info(
                "Kernel settings loaded. workbook="
                + GetWorkbookFullName(workbook)
                + ", systemRoot="
                + (systemRoot ?? string.Empty)
                + ", defaultRoot="
                + (defaultRoot ?? string.Empty)
                + ", nameRuleA="
                + (nameRuleA ?? string.Empty)
                + ", nameRuleB="
                + (nameRuleB ?? string.Empty));

            return new KernelSettingsState
            {
                SystemRoot = systemRoot,
                DefaultRoot = defaultRoot,
                NameRuleA = nameRuleA,
                NameRuleB = nameRuleB
            };
        }

        internal bool SaveNameRuleA(string ruleA)
        {
            Excel.Workbook workbook = ResolveWorkbookForBoundHomeMutation("SaveNameRuleA");
            if (workbook == null)
            {
                return false;
            }

            _kernelWorkbookSettingsService.SaveNameRuleA(workbook, ruleA);
            workbook.Save();
            return true;
        }

        internal bool SaveNameRuleB(string ruleB)
        {
            Excel.Workbook workbook = ResolveWorkbookForBoundHomeMutation("SaveNameRuleB");
            if (workbook == null)
            {
                return false;
            }

            _kernelWorkbookSettingsService.SaveNameRuleB(workbook, ruleB);
            workbook.Save();
            return true;
        }

        internal string SelectAndSaveDefaultRoot()
        {
            Excel.Workbook workbook = ResolveWorkbookForBoundHomeMutation("SelectAndSaveDefaultRoot");
            if (workbook == null)
            {
                return null;
            }

            string selectedPath = _kernelWorkbookSettingsService.SelectDefaultRoot(workbook);
            if (string.IsNullOrWhiteSpace(selectedPath))
            {
                return null;
            }

            workbook.Save();
            return selectedPath;
        }

        internal Excel.Workbook ResolveWorkbookForHomeDisplayOrClose(string operationName)
        {
            Excel.Workbook workbook;
            KernelHomeBindingStatus bindingStatus = ResolveHomeBindingStatus(operationName, out workbook);
            if (bindingStatus == KernelHomeBindingStatus.Valid)
            {
                return workbook;
            }

            _logger.Info(
                "Kernel HOME display/close operation skipped because valid binding was not available. operation="
                + (operationName ?? string.Empty)
                + ", bindingStatus="
                + bindingStatus.ToString());
            return null;
        }

        internal bool HasOtherVisibleWorkbook(Excel.Workbook workbookToIgnore)
        {
            return _kernelWorkbookStateService.HasOtherVisibleWorkbook(workbookToIgnore);
        }

        internal bool HasOtherWorkbook(Excel.Workbook workbookToIgnore)
        {
            return _kernelWorkbookStateService.HasOtherWorkbook(workbookToIgnore);
        }

        internal bool HasVisibleNonKernelWorkbook()
        {
            return _kernelWorkbookStateService.HasVisibleNonKernelWorkbook();
        }

        internal string GetWorkbookFullName(Excel.Workbook workbook)
        {
            if (_excelInteropService != null)
            {
                return _excelInteropService.GetWorkbookFullName(workbook);
            }

            try
            {
                return workbook == null ? string.Empty : workbook.FullName ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        internal string GetWorkbookName(Excel.Workbook workbook)
        {
            if (_excelInteropService != null)
            {
                return _excelInteropService.GetWorkbookName(workbook);
            }

            try
            {
                return workbook == null ? string.Empty : workbook.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private KernelHomeBinding CreateHomeBinding(WorkbookContext context)
        {
            if (context == null)
            {
                return null;
            }

            string normalizedContextSystemRoot = _pathCompatibilityService.NormalizePath(context.SystemRoot);
            if (string.IsNullOrWhiteSpace(normalizedContextSystemRoot))
            {
                return null;
            }

            return CreateHomeBinding(context.Workbook, normalizedContextSystemRoot);
        }

        private KernelHomeBinding CreateHomeBinding(Excel.Workbook workbook)
        {
            return CreateHomeBinding(workbook, systemRoot: null);
        }

        private KernelHomeBinding CreateHomeBinding(Excel.Workbook workbook, string systemRoot)
        {
            if (!IsKernelWorkbook(workbook))
            {
                return null;
            }

            string workbookSystemRoot = GetWorkbookSystemRootForHomeBinding(workbook);
            if (string.IsNullOrWhiteSpace(workbookSystemRoot))
            {
                return null;
            }

            string normalizedSystemRoot = _pathCompatibilityService.NormalizePath(systemRoot);
            if (!string.IsNullOrWhiteSpace(normalizedSystemRoot)
                && !string.Equals(workbookSystemRoot, normalizedSystemRoot, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            return new KernelHomeBinding(workbook, string.IsNullOrWhiteSpace(normalizedSystemRoot) ? workbookSystemRoot : normalizedSystemRoot);
        }

        private KernelHomeBindingStatus ResolveHomeBindingStatus(string operationName, out Excel.Workbook workbook)
        {
            workbook = null;
            if (_homeBinding == null)
            {
                return KernelHomeBindingStatus.None;
            }

            Excel.Workbook boundWorkbook = _homeBinding.Workbook;
            if (!IsKernelWorkbook(boundWorkbook))
            {
                _logger.Warn(
                    "Kernel HOME binding became invalid because workbook was not kernel. operation="
                    + (operationName ?? string.Empty)
                    + ", workbook="
                    + GetWorkbookFullName(boundWorkbook));
                return KernelHomeBindingStatus.Invalid;
            }

            string currentSystemRoot = GetWorkbookSystemRootForHomeBinding(boundWorkbook);
            if (string.IsNullOrWhiteSpace(currentSystemRoot)
                || !string.Equals(currentSystemRoot, _homeBinding.SystemRoot, StringComparison.OrdinalIgnoreCase))
            {
                _logger.Warn(
                    "Kernel HOME binding became invalid because system root mismatched. operation="
                    + (operationName ?? string.Empty)
                    + ", workbook="
                    + GetWorkbookFullName(boundWorkbook)
                    + ", boundSystemRoot="
                    + (_homeBinding.SystemRoot ?? string.Empty)
                    + ", currentSystemRoot="
                    + (currentSystemRoot ?? string.Empty));
                return KernelHomeBindingStatus.Invalid;
            }

            workbook = boundWorkbook;
            return KernelHomeBindingStatus.Valid;
        }

        private Excel.Workbook ResolveWorkbookForBoundHomeMutation(string operationName)
        {
            Excel.Workbook workbook;
            KernelHomeBindingStatus bindingStatus = ResolveHomeBindingStatus(operationName, out workbook);
            if (bindingStatus != KernelHomeBindingStatus.Valid)
            {
                _logger.Warn(
                    "Kernel HOME mutation failed closed because valid binding was not available. operation="
                    + (operationName ?? string.Empty)
                    + ", bindingStatus="
                    + bindingStatus.ToString());
                return null;
            }

            return workbook;
        }

        private string GetWorkbookSystemRootForHomeBinding(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return string.Empty;
            }

            if (_excelInteropService != null)
            {
                return _pathCompatibilityService.NormalizePath(
                    KernelWorkbookResolver.GetSystemRootFromWorkbook(
                        workbook,
                        _excelInteropService,
                        _pathCompatibilityService,
                        _logger,
                        IsKernelWorkbook));
            }

            try
            {
                if (workbook.CustomDocumentProperties is IDictionary<string, string> properties
                    && properties.TryGetValue("SYSTEM_ROOT", out string systemRoot))
                {
                    return _pathCompatibilityService.NormalizePath(systemRoot);
                }
            }
            catch
            {
            }

            return string.Empty;
        }

        private sealed class KernelHomeBinding
        {
            internal KernelHomeBinding(Excel.Workbook workbook, string systemRoot)
            {
                Workbook = workbook;
                SystemRoot = systemRoot ?? string.Empty;
            }

            internal Excel.Workbook Workbook { get; }

            internal string SystemRoot { get; }
        }

        private enum KernelHomeBindingStatus
        {
            None = 0,
            Valid = 1,
            Invalid = 2
        }
    }
}
