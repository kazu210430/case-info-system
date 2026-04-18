using System;
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
        private string _loadedModePath;
        private DocumentExecutionMode _currentMode;

        /// <summary>
        internal DocumentExecutionModeService(Logger logger, ExcelInteropService excelInteropService)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _loadedModePath = string.Empty;
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
            if (string.Equals(_loadedModePath, modePath, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            _loadedModePath = modePath;
            _currentMode = LoadMode(modePath);
        }

        /// <summary>
        private string ResolveModeFilePath()
        {
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

        /// <summary>
        private DocumentExecutionMode LoadMode(string modePath)
        {
            if (!File.Exists(modePath))
            {
                _logger.Info("Document execution mode file was not found. path=" + modePath + ", mode=Disabled");
                return DocumentExecutionMode.Disabled;
            }

            try
            {
                string rawValue = File.ReadLines(modePath)
                    .Select(line => (line ?? string.Empty).Trim())
                    .FirstOrDefault(line => line.Length > 0 && !line.StartsWith("#", StringComparison.Ordinal));
                if (string.Equals(rawValue, PilotOnlyModeName, StringComparison.OrdinalIgnoreCase))
                {
                    _logger.Info("Document execution mode loaded. path=" + modePath + ", mode=PilotOnly");
                    return DocumentExecutionMode.PilotOnly;
                }

                if (string.Equals(rawValue, AllowlistedOnlyModeName, StringComparison.OrdinalIgnoreCase))
                {
                    _logger.Info("Document execution mode loaded. path=" + modePath + ", mode=AllowlistedOnly");
                    return DocumentExecutionMode.AllowlistedOnly;
                }

                if (string.Equals(rawValue, DisabledModeName, StringComparison.OrdinalIgnoreCase) || string.IsNullOrWhiteSpace(rawValue))
                {
                    _logger.Info("Document execution mode loaded. path=" + modePath + ", mode=Disabled");
                    return DocumentExecutionMode.Disabled;
                }

                _logger.Warn("Document execution mode file contained invalid value. path=" + modePath + ", value=" + rawValue + ", fallback=Disabled");
                return DocumentExecutionMode.Disabled;
            }
            catch (Exception ex)
            {
                _logger.Error("Document execution mode load failed. path=" + modePath, ex);
                return DocumentExecutionMode.Disabled;
            }
        }
    }
}
