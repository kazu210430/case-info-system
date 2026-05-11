using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal static class CaseListMappingKeyNormalizer
	{
		internal static IReadOnlyList<CaseListMappingDefinition> NormalizeSourceFieldKeys(IReadOnlyList<CaseListMappingDefinition> mappings)
		{
			if (mappings == null || mappings.Count == 0)
			{
				return Array.Empty<CaseListMappingDefinition>();
			}

			List<CaseListMappingDefinition> result = new List<CaseListMappingDefinition>(mappings.Count);
			foreach (CaseListMappingDefinition mapping in mappings)
			{
				if (mapping == null)
				{
					continue;
				}

				result.Add(new CaseListMappingDefinition
				{
					MappingType = mapping.MappingType,
					SourceFieldKey = FieldKeyRenameMap.NormalizeToCurrent(mapping.SourceFieldKey),
					TargetHeaderName = mapping.TargetHeaderName,
					DataType = mapping.DataType,
					NormalizeRule = mapping.NormalizeRule
				});
			}

			return result;
		}
	}
}
