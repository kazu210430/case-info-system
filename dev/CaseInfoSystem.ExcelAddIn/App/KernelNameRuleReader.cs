using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Excel = Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
    internal sealed class KernelNameRuleReader
    {
        private const string SystemRootPropertyName = "SYSTEM_ROOT";
        private const string NameRuleAPropertyName = "NAME_RULE_A";
        private const string NameRuleBPropertyName = "NAME_RULE_B";
        private const string DefaultNameRuleA = "YYYY";
        private const string DefaultNameRuleB = "DOC";
        private static readonly XNamespace CustomPropertiesNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";

        private readonly ExcelInteropService _excelInteropService;
        private readonly PathCompatibilityService _pathCompatibilityService;
        private readonly Logger _logger;

        internal KernelNameRuleReader(
            ExcelInteropService excelInteropService,
            PathCompatibilityService pathCompatibilityService,
            Logger logger)
        {
            _excelInteropService = excelInteropService ?? throw new ArgumentNullException(nameof(excelInteropService));
            _pathCompatibilityService = pathCompatibilityService ?? throw new ArgumentNullException(nameof(pathCompatibilityService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        internal bool TryReadForCaseWorkbook(Excel.Workbook caseWorkbook, out string ruleA, out string ruleB)
        {
            ruleA = KernelNamingService.NormalizeNameRuleA(DefaultNameRuleA);
            ruleB = KernelNamingService.NormalizeNameRuleB(DefaultNameRuleB);

            string kernelPath = ResolveKernelWorkbookPath(caseWorkbook);
            if (string.IsNullOrWhiteSpace(kernelPath) || !_pathCompatibilityService.FileExistsSafe(kernelPath))
            {
                return false;
            }

            if (!TryGetKernelNameRules(kernelPath, out string rawRuleA, out string rawRuleB))
            {
                return false;
            }

            ruleA = KernelNamingService.NormalizeNameRuleA(rawRuleA);
            ruleB = KernelNamingService.NormalizeNameRuleB(rawRuleB);
            return true;
        }

        private string ResolveKernelWorkbookPath(Excel.Workbook caseWorkbook)
        {
            string systemRoot = _pathCompatibilityService.NormalizePath(
                _excelInteropService.TryGetDocumentProperty(caseWorkbook, SystemRootPropertyName));
            if (systemRoot.Length == 0)
            {
                return string.Empty;
            }

            return WorkbookFileNameResolver.ResolveExistingKernelWorkbookPath(systemRoot, _pathCompatibilityService);
        }

        private bool TryGetKernelNameRules(string kernelPath, out string ruleA, out string ruleB)
        {
            ruleA = DefaultNameRuleA;
            ruleB = DefaultNameRuleB;

            Excel.Workbook openKernelWorkbook = _excelInteropService.FindOpenWorkbook(kernelPath);
            if (openKernelWorkbook != null)
            {
                ruleA = _excelInteropService.TryGetDocumentProperty(openKernelWorkbook, NameRuleAPropertyName);
                ruleB = _excelInteropService.TryGetDocumentProperty(openKernelWorkbook, NameRuleBPropertyName);
                return true;
            }

            return TryReadKernelNameRulesFromPackage(kernelPath, out ruleA, out ruleB);
        }

        private bool TryReadKernelNameRulesFromPackage(string kernelPath, out string ruleA, out string ruleB)
        {
            ruleA = DefaultNameRuleA;
            ruleB = DefaultNameRuleB;

            try
            {
                using (ZipArchive archive = ZipFile.OpenRead(kernelPath))
                {
                    ZipArchiveEntry customXmlEntry = archive.GetEntry("docProps/custom.xml");
                    if (customXmlEntry == null)
                    {
                        return false;
                    }

                    using (Stream stream = customXmlEntry.Open())
                    {
                        XDocument document = XDocument.Load(stream);
                        foreach (XElement propertyElement in document.Root == null
                            ? Array.Empty<XElement>()
                            : document.Root.Elements(CustomPropertiesNamespace + "property"))
                        {
                            XAttribute nameAttribute = propertyElement.Attribute("name");
                            if (nameAttribute == null)
                            {
                                continue;
                            }

                            XElement valueElement = propertyElement.Elements().FirstOrDefault();
                            if (valueElement == null)
                            {
                                continue;
                            }

                            string propertyValue = valueElement.Value ?? string.Empty;
                            if (string.Equals(nameAttribute.Value, NameRuleAPropertyName, StringComparison.OrdinalIgnoreCase))
                            {
                                ruleA = propertyValue;
                            }
                            else if (string.Equals(nameAttribute.Value, NameRuleBPropertyName, StringComparison.OrdinalIgnoreCase))
                            {
                                ruleB = propertyValue;
                            }
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.Error("Kernel package name-rule read failed. path=" + kernelPath, ex);
                return false;
            }
        }
    }
}
