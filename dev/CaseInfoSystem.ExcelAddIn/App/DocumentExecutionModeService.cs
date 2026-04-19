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
        private const string PilotOnlyModeName = "PilotOnly";
        private const string AllowlistedOnlyModeName = "AllowlistedOnly";
        private const string DefaultSystemRootFolderName = "\u6848\u4EF6\u60C5\u5831System";
        private const string AddInsFolderName = "Addins";
        private const string RuntimeAddInFolderName = "CaseInfoSystem.ExcelAddIn";
        private const string SystemRootPropertyName = "SYSTEM_ROOT";

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
        internal DocumentExecutionMode GetMode()
        {
            EnsureLoaded();
            return _currentMode;
        }

        /// <summary>
        internal bool CanAttemptVstoExecution()
        {
            DocumentExecutionMode currentMode = GetMode();
            return currentMode == DocumentExecutionMode.PilotOnly
                || currentMode == DocumentExecutionMode.AllowlistedOnly;
        }

        /// <summary>
        private void EnsureLoaded()
        {
            string modePath = ResolveModeFilePath();
            DateTime? currentLastWriteTimeUtc = TryGetModeFileLastWriteTimeUtc(modePath);
            if (string.Equals(_loadedModePath, modePath, StringComparison.OrdinalIgnoreCase)
                && Nullable.Equals(_loadedModeLastWriteTimeUtc, currentLastWriteTimeUtc))
            {
                return;
            }

            DocumentExecutionMode loadedMode;
            if (!TryLoadMode(modePath, out loadedMode))
            {
                return;
            }

            _loadedModePath = modePath;
            _loadedModeLastWriteTimeUtc = currentLastWriteTimeUtc;
            _currentMode = loadedMode;
        }

        /// <summary>
        private string ResolveModeFilePath()
        {
            if (_testHooks != null && _testHooks.ResolveModeFilePath != null)
            {
                return _testHooks.ResolveModeFilePath() ?? string.Empty;
            }

            string runtimeDirectory = ResolveRuntimeAddInDirectory();
            if (!string.IsNullOrWhiteSpace(runtimeDirectory))
            {
                return Path.Combine(runtimeDirectory, ModeFileName);
            }

            string assemblyDirectory = Path.GetDirectoryName(typeof(DocumentExecutionModeService).Assembly.Location) ?? string.Empty;
            return Path.Combine(assemblyDirectory, ModeFileName);
        }

        /// <summary>
        private string ResolveRuntimeAddInDirectory()
        {
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
                    return candidateDirectory;
                }
            }

            string documentsRoot = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (string.IsNullOrWhiteSpace(documentsRoot))
            {
                return string.Empty;
            }

            string fallbackDirectory = Path.Combine(
                documentsRoot,
                DefaultSystemRootFolderName,
                AddInsFolderName,
                RuntimeAddInFolderName);
            return Directory.Exists(fallbackDirectory) ? fallbackDirectory : string.Empty;
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

        private bool TryLoadMode(string modePath, out DocumentExecutionMode loadedMode)
        {
            loadedMode = _currentMode;
            if (!ModeFileExists(modePath))
            {
                _logger.Info("Document execution mode file was not found. path=" + modePath + ", mode=Disabled");
                loadedMode = DocumentExecutionMode.Disabled;
                return true;
            }

            try
            {
                string rawValue = ReadModeFileLines(modePath)
                    .Select(line => (line ?? string.Empty).Trim())
                    .FirstOrDefault(line => line.Length > 0 && !line.StartsWith("#", StringComparison.Ordinal));
                if (string.Equals(rawValue, PilotOnlyModeName, StringComparison.OrdinalIgnoreCase))
                {
                    _logger.Info("Document execution mode loaded. path=" + modePath + ", mode=PilotOnly");
                    loadedMode = DocumentExecutionMode.PilotOnly;
                    return true;
                }

                if (string.Equals(rawValue, AllowlistedOnlyModeName, StringComparison.OrdinalIgnoreCase))
                {
                    _logger.Info("Document execution mode loaded. path=" + modePath + ", mode=AllowlistedOnly");
                    loadedMode = DocumentExecutionMode.AllowlistedOnly;
                    return true;
                }

                if (string.Equals(rawValue, DisabledModeName, StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(rawValue))
                {
                    _logger.Info("Document execution mode loaded. path=" + modePath + ", mode=Disabled");
                    loadedMode = DocumentExecutionMode.Disabled;
                    return true;
                }

                _logger.Warn("Document execution mode file contained invalid value. path=" + modePath + ", value=" + rawValue + ", fallback=Disabled");
                loadedMode = DocumentExecutionMode.Disabled;
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error("Document execution mode load failed. path=" + modePath, ex);
                return false;
            }
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

        internal sealed class DocumentExecutionModeServiceTestHooks
        {
            internal Func<string> ResolveModeFilePath { get; set; }

            internal Func<string, bool> ModeFileExists { get; set; }

            internal Func<string, IEnumerable<string>> ReadModeFileLines { get; set; }

            internal Func<string, DateTime?> GetModeFileLastWriteTimeUtc { get; set; }
        }
    }
}
