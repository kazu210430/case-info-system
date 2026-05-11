using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public sealed class CaseListMappingKeyNormalizerTests
	{
		[Fact]
		public void NormalizeSourceFieldKeys_WhenLegacyLawyerMappingExists_UsesCurrentDisplayNameKey()
		{
			IReadOnlyList<CaseListMappingDefinition> normalized = CaseListMappingKeyNormalizer.NormalizeSourceFieldKeys(
				new[]
				{
					Mapping(FieldKeyRenameMap.LegacyLawyerKey, "担当")
				});

			CaseListMappingDefinition mapping = Assert.Single(normalized);
			Assert.Equal(FieldKeyRenameMap.CurrentLawyerKey, mapping.SourceFieldKey);
			Assert.DoesNotContain(normalized, item => string.Equals(item.SourceFieldKey, FieldKeyRenameMap.LegacyLawyerKey, StringComparison.OrdinalIgnoreCase));
		}

		[Fact]
		public void NormalizeSourceFieldKeys_WhenInventoryUsesCurrentKey_MappingCanMatchWithoutAddingLegacyInventory()
		{
			IReadOnlyDictionary<string, CaseListFieldDefinition> fieldInventory =
				new Dictionary<string, CaseListFieldDefinition>(StringComparer.OrdinalIgnoreCase)
				{
					[FieldKeyRenameMap.CurrentLawyerKey] = new CaseListFieldDefinition
					{
						FieldKey = FieldKeyRenameMap.CurrentLawyerKey
					}
				};

			CaseListMappingDefinition mapping = Assert.Single(CaseListMappingKeyNormalizer.NormalizeSourceFieldKeys(
				new[]
				{
					Mapping(FieldKeyRenameMap.LegacyLawyerKey, "担当")
				}));

			Assert.True(fieldInventory.ContainsKey(mapping.SourceFieldKey));
			Assert.False(fieldInventory.ContainsKey(FieldKeyRenameMap.LegacyLawyerKey));
		}

		[Fact]
		public void NormalizeSourceFieldKeys_WhenUnknownKeyExists_DoesNotMakeItValid()
		{
			CaseListMappingDefinition mapping = Assert.Single(CaseListMappingKeyNormalizer.NormalizeSourceFieldKeys(
				new[]
				{
					Mapping("未知キー", "担当")
				}));

			Assert.Equal("未知キー", mapping.SourceFieldKey);
		}

		private static CaseListMappingDefinition Mapping(string sourceFieldKey, string targetHeaderName)
		{
			return new CaseListMappingDefinition
			{
				MappingType = "Direct",
				SourceFieldKey = sourceFieldKey,
				TargetHeaderName = targetHeaderName,
				DataType = "Text",
				NormalizeRule = "Trim"
			};
		}
	}
}
