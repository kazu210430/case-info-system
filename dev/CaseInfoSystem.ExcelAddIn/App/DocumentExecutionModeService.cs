using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;

namespace CaseInfoSystem.ExcelAddIn.App
{
    /// <summary>
    internal sealed class DocumentExecutionModeService
    {
        private const string ModeFileName = "DocumentExecutionMode.txt";
        private const string DisabledModeName = "Disabled";
        private const string WarmupEnabledProfileAModeName = "WarmupEnabledProfileA";
        private const string WarmupEnabledProfileBModeName = "WarmupEnabledProfileB";
        private const string LegacyPilotOnlyModeName = "PilotOnly";
        private const string LegacyAllowlistedOnlyModeName = "AllowlistedOnly";
        private const string DefaultDocumentsSystemRootFolderName = "\u6848\u4EF6\u60C5\u5831System";
        private const string AddInsFolderName = "Addins";
        private const string RuntimeAddInFolderName = "CaseInfoSystem.ExcelAddIn";
        private const string SystemRootPropertyName = "SYSTEM_ROOT";
        private const string TestOverrideResolutionSource = "TestOverride";
        private const string OpenWorkbookRuntimeResolutionSource = "OpenWorkbookSystemRootRuntime";
        private const string DefaultDocumentsRuntimeFallbackResolutionSource = "DefaultDocumentsRuntimeFallback";
        private const string AssemblyBootstrapFallbackResolutionSource = "AssemblyBootstrapFallback";

        private readonly Logger _logger;
        private readonly ExcelInteropService _excelInteropService;
        private readonly DocumentExecutionModeServiceTestHooks _testHooks;
        private string _loadedModePath;
        private DateTime? _loadedModeLastWriteTimeUtc;
        private DocumentExecutionMode _currentMode;

        /// <summary>
        internal DocumentExecutionModeService(Logger logger, ExcelInteropService excelInteropService)
            : this(logger, excelInteropService, testHooks: null)
        {
        }

        internal DocumentExecutionModeService(Logger logger, ExcelInteropService excelInteropService, DocumentExecutionModeServiceTestHooks testHooks)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _testHooks = testHooks;
            _loadedModePath = string.Empty;
            _loadedModeLastWriteTimeUtc = null;
            _currentMode = DocumentExecutionMode.Disabled;
        }

        /// <summary>
        internal DocumentExecutionMode GetConfiguredMode()
        {
            EnsureLoaded();
            return _currentMode;
        }

        /// <summary>
        internal bool IsWordWarmupEnabled()
        {
            DocumentExecutionMode currentMode = GetConfiguredMode();
            return currentMode == DocumentExecutionMode.WarmupEnabledProfileA
                || currentMode == DocumentExecutionMode.WarmupEnabledProfileB;
        }

        /// <summary>
        private void EnsureLoaded()
        {
            DocumentExecutionModeFileLocation modeLocation = ResolveModeFileLocation();
            DateTime? currentLastWriteTimeUtc = TryGetModeFileLastWriteTimeUtc(modeLocation.FilePath);
            if (string.Equals(_loadedModePath, modeLocation.FilePath, StringComparison.OrdinalIgnoreCase)
                && Nullable.Equals(_loadedModeLastWriteTimeUtc, currentLastWriteTimeUtc))
            {
                return;
            }

            DocumentExecutionMode loadedMode;
            if (!TryLoadMode(modeLocation, out loadedMode))
            {
                return;
            }

            _loadedModePath = modeLocation.FilePath;
            _loadedModeLastWriteTimeUtc = currentLastWriteTimeUtc;
            _currentMode = loadedMode;
        }

        /// <summary>
        private DocumentExecutionModeFileLocation ResolveModeFileLocation()
        {
            if (_testHooks != null && _testHooks.ResolveModeFilePath != null)
            {
                return new DocumentExecutionModeFileLocation(
                    _testHooks.ResolveModeFilePath() ?? string.Empty,
                    TestOverrideResolutionSource);
            }

            string runtimeDirectory;
            if (TryResolveRuntimeAddInDirectoryFromOpenWorkbooks(out runtimeDirectory))
            {
                return CreateModeFileLocation(runtimeDirectory, OpenWorkbookRuntimeResolutionSource);
            }

            if (TryResolveRuntimeAddInDirectoryFromDefaultDocumentsRoot(out runtimeDirectory))
            {
                return CreateModeFileLocation(runtimeDirectory, DefaultDocumentsRuntimeFallbackResolutionSource);
            }

            string assemblyDirectory = Path.GetDirectoryName(typeof(DocumentExecutionModeService).Assembly.Location) ?? string.Empty;
            return CreateModeFileLocation(assemblyDirectory, AssemblyBootstrapFallbackResolutionSource);
        }

        /// <summary>
        private bool TryResolveRuntimeAddInDirectoryFromOpenWorkbooks(out string runtimeDirectory)
        {
            runtimeDirectory = string.Empty;
            foreach (var workbook in _excelInteropService.GetOpenWorkbooks())
            {
                string systemRoot = (_excelInteropService.TryGetDocumentProperty(workbook, SystemRootPropertyName) ?? string.Empty).Trim();
                if (systemRoot.Length == 0)
                {
                    continue;
                }

                string candidateDirectory = Path.Combine(
                    systemRoot,
                    AddInsFolderName,
                    RuntimeAddInFolderName);
                if (Directory.Exists(candidateDirectory))
                {
                    runtimeDirectory = candidateDirectory;
                    return true;
                }
            }

            return false;
        }

        private bool TryResolveRuntimeAddInDirectoryFromDefaultDocumentsRoot(out string runtimeDirectory)
        {
            runtimeDirectory = string.Empty;
            string documentsRoot = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (string.IsNullOrWhiteSpace(documentsRoot))
            {
                return false;
            }

            string fallbackDirectory = Path.Combine(
                documentsRoot,
                DefaultDocumentsSystemRootFolderName,
                AddInsFolderName,
                RuntimeAddInFolderName);
            if (!Directory.Exists(fallbackDirectory))
            {
                return false;
            }

            runtimeDirectory = fallbackDirectory;
            return true;
        }

        private DateTime? TryGetModeFileLastWriteTimeUtc(string modePath)
        {
            try
            {
                if (_testHooks != null && _testHooks.GetModeFileLastWriteTimeUtc != null)
                {
                    return _testHooks.GetModeFileLastWriteTimeUtc(modePath);
                }

                if (!File.Exists(modePath))
                {
                    return null;
                }

                return File.GetLastWriteTimeUtc(modePath);
            }
            catch (Exception ex)
            {
                _logger.Error("Document execution mode timestamp read failed. path=" + modePath, ex);
                return null;
            }
        }

        private bool TryLoadMode(DocumentExecutionModeFileLocation modeLocation, out DocumentExecutionMode loadedMode)
        {
            string modePath = modeLocation.FilePath;
            loadedMode = _currentMode;
            if (!ModeFileExists(modePath))
            {
                _logger.Info(
                    "Document execution mode file was not found. path=" + modePath
                    + ", source=" + modeLocation.ResolutionSource
                    + ", mode=Disabled");
                loadedMode = DocumentExecutionMode.Disabled;
                return true;
            }

            try
            {
                string rawValue = ReadModeFileLines(modePath)
                    .Select(line => (line ?? string.Empty).Trim())
                    .FirstOrDefault(line => line.Length > 0 && !line.StartsWith("#", StringComparison.Ordinal));
                if (MatchesAny(rawValue, WarmupEnabledProfileAModeName, LegacyPilotOnlyModeName))
                {
                    loadedMode = DocumentExecutionMode.WarmupEnabledProfileA;
                    LogLoadedMode(modeLocation, rawValue, loadedMode);
                    return true;
                }

                if (MatchesAny(rawValue, WarmupEnabledProfileBModeName, LegacyAllowlistedOnlyModeName))
                {
                    loadedMode = DocumentExecutionMode.WarmupEnabledProfileB;
                    LogLoadedMode(modeLocation, rawValue, loadedMode);
                    return true;
                }

                if (string.Equals(rawValue, DisabledModeName, StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(rawValue))
                {
                    loadedMode = DocumentExecutionMode.Disabled;
                    LogLoadedMode(modeLocation, rawValue, loadedMode);
                    return true;
                }

                _logger.Warn(
                    "Document execution mode file contained invalid value. path=" + modePath
                    + ", source=" + modeLocation.ResolutionSource
                    + ", value=" + rawValue
                    + ", fallback=Disabled");
                loadedMode = DocumentExecutionMode.Disabled;
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(
                    "Document execution mode load failed. path=" + modePath
                    + ", source=" + modeLocation.ResolutionSource,
                    ex);
                return false;
            }
        }

        private DocumentExecutionModeFileLocation CreateModeFileLocation(string directoryPath, string resolutionSource)
        {
            return new DocumentExecutionModeFileLocation(
                Path.Combine(directoryPath ?? string.Empty, ModeFileName),
                resolutionSource);
        }

        private void LogLoadedMode(DocumentExecutionModeFileLocation modeLocation, string rawValue, DocumentExecutionMode loadedMode)
        {
            string message =
                "Document execution mode loaded. path=" + modeLocation.FilePath
                + ", source=" + modeLocation.ResolutionSource
                + ", mode=" + loadedMode.ToString();
            if (!string.IsNullOrWhiteSpace(rawValue)
                && !string.Equals(rawValue, loadedMode.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                message += ", rawValue=" + rawValue;
            }

            _logger.Info(message);
        }

        private bool MatchesAny(string rawValue, params string[] acceptedValues)
        {
            foreach (string acceptedValue in acceptedValues)
            {
                if (string.Equals(rawValue, acceptedValue, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private bool ModeFileExists(string modePath)
        {
            return _testHooks != null && _testHooks.ModeFileExists != null
                ? _testHooks.ModeFileExists(modePath)
                : File.Exists(modePath);
        }

        private IEnumerable<string> ReadModeFileLines(string modePath)
        {
            return _testHooks != null && _testHooks.ReadModeFileLines != null
                ? _testHooks.ReadModeFileLines(modePath)
                : File.ReadLines(modePath);
        }

        private sealed class DocumentExecutionModeFileLocation
        {
            internal DocumentExecutionModeFileLocation(string filePath, string resolutionSource)
            {
                FilePath = filePath ?? string.Empty;
                ResolutionSource = resolutionSource ?? string.Empty;
            }

            internal string FilePath { get; private set; }

            internal string ResolutionSource { get; private set; }
        }

        internal sealed class DocumentExecutionModeServiceTestHooks
        {
            internal Func<string> ResolveModeFilePath { get; set; }

            internal Func<string, bool> ModeFileExists { get; set; }

            internal Func<string, IEnumerable<string>> ReadModeFileLines { get; set; }

            internal Func<string, DateTime?> GetModeFileLastWriteTimeUtc { get; set; }
        }
    }
}
