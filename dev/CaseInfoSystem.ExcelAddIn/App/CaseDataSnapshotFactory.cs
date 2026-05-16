using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.Infrastructure;
using Microsoft.Office.Interop.Excel;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal sealed class CaseDataSnapshotFactory
	{
		private const string HomeSheetCodeName = "shHOME";

		private const string CustomerNameFieldKey = "\u9867\u5BA2_\u540D\u524D";

		private readonly ExcelInteropService _excelInteropService;

		private readonly KernelWorkbookResolverService _kernelWorkbookResolverService;

		private readonly CaseListFieldDefinitionRepository _fieldDefinitionRepository;

		private readonly Logger _logger;

		internal CaseDataSnapshotFactory (ExcelInteropService excelInteropService, KernelWorkbookResolverService kernelWorkbookResolverService, CaseListFieldDefinitionRepository fieldDefinitionRepository, Logger logger)
		{
			_excelInteropService = excelInteropService ?? throw new ArgumentNullException ("excelInteropService");
			_kernelWorkbookResolverService = kernelWorkbookResolverService ?? throw new ArgumentNullException ("kernelWorkbookResolverService");
			_fieldDefinitionRepository = fieldDefinitionRepository ?? throw new ArgumentNullException ("fieldDefinitionRepository");
			_logger = logger ?? throw new ArgumentNullException ("logger");
		}

		internal CaseDataSnapshot Create (Workbook caseWorkbook)
		{
			KernelWorkbookAccessResult kernelAccess = _kernelWorkbookResolverService.ResolveOrOpenReadOnly (caseWorkbook);
			Workbook workbook = kernelAccess.Workbook;
			try {
				return Create (caseWorkbook, workbook);
			} finally {
				kernelAccess.CloseIfOwned (
					"CaseDataSnapshotFactory.CloseOpenedKernelWorkbook",
					suppressEventsDuringClose: true);
			}
		}

		internal CaseDataSnapshot Create (Workbook caseWorkbook, Workbook kernelWorkbook)
		{
			CaseDataSnapshot result = new CaseDataSnapshot {
				Values = new Dictionary<string, string> (StringComparer.OrdinalIgnoreCase),
				CustomerName = string.Empty
			};
			if (caseWorkbook == null || kernelWorkbook == null) {
				return result;
			}
			Worksheet worksheet = _excelInteropService.FindWorksheetByCodeName (caseWorkbook, "shHOME");
			if (worksheet == null) {
				return result;
			}
			IReadOnlyDictionary<string, CaseListFieldDefinition> readOnlyDictionary = _fieldDefinitionRepository.LoadDefinitions (kernelWorkbook);
			if (readOnlyDictionary == null || readOnlyDictionary.Count == 0) {
				return result;
			}
			IReadOnlyDictionary<string, string> readOnlyDictionary2 = _excelInteropService.ReadFieldValuesFromDefinitions (worksheet, readOnlyDictionary.Values);
			if (!readOnlyDictionary2.TryGetValue (CustomerNameFieldKey, out var value)) {
				value = string.Empty;
			}
			_logger.Info ("Case data snapshot resolved by field definitions. fields=" + readOnlyDictionary2.Count);
			return new CaseDataSnapshot {
				Values = readOnlyDictionary2,
				CustomerName = (value ?? string.Empty).Trim ()
			};
		}
	}
}
