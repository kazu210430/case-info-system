using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;

namespace CaseInfoSystem.ExcelAddIn.App
{
	internal static class CaseListManagedKeyNormalizer
	{
		internal static IReadOnlyList<CaseListMappingDefinition> NormalizeMappingKeys(IReadOnlyList<CaseListMappingDefinition> mappings)
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
					SourceFieldKey = NormalizeFieldKey(mapping.SourceFieldKey),
					TargetHeaderName = NormalizeHeaderName(mapping.TargetHeaderName),
					DataType = mapping.DataType,
					NormalizeRule = mapping.NormalizeRule
				});
			}

			return result;
		}

		internal static string NormalizeFieldKey(string fieldKey)
		{
			return FieldKeyRenameMap.NormalizeToCurrent(fieldKey);
		}

		internal static string NormalizeHeaderName(string headerName)
		{
			return FieldKeyRenameMap.NormalizeToCurrent(headerName);
		}
	}
}
