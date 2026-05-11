using System;
using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.App;
using CaseInfoSystem.ExcelAddIn.Domain;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public sealed class CaseListManagedKeyNormalizerTests
	{
		[Fact]
		public void NormalizeMappingKeys_WhenLegacyLawyerMappingExists_UsesCurrentDisplayNameKey()
		{
			IReadOnlyList<CaseListMappingDefinition> normalized = CaseListManagedKeyNormalizer.NormalizeMappingKeys(
				new[]
				{
					Mapping(FieldKeyRenameMap.LegacyLawyerKey, "担当")
				});

			CaseListMappingDefinition mapping = Assert.Single(normalized);
			Assert.Equal(FieldKeyRenameMap.CurrentLawyerKey, mapping.SourceFieldKey);
			Assert.DoesNotContain(normalized, item => string.Equals(item.SourceFieldKey, FieldKeyRenameMap.LegacyLawyerKey, StringComparison.OrdinalIgnoreCase));
		}

		[Fact]
		public void NormalizeMappingKeys_WhenLegacyLawyerHeaderExists_UsesCurrentDisplayNameHeader()
		{
			CaseListMappingDefinition mapping = Assert.Single(CaseListManagedKeyNormalizer.NormalizeMappingKeys(
				new[]
				{
					Mapping(FieldKeyRenameMap.CurrentLawyerKey, FieldKeyRenameMap.LegacyLawyerKey)
				}));

			Assert.Equal(FieldKeyRenameMap.CurrentLawyerKey, mapping.TargetHeaderName);
		}

		[Fact]
		public void NormalizeHeaderName_WhenLegacyLawyerHeaderIsManaged_UsesCurrentDisplayNameHeader()
		{
			string normalizedHeader = CaseListManagedKeyNormalizer.NormalizeHeaderName(FieldKeyRenameMap.LegacyLawyerKey);

			Assert.Equal(FieldKeyRenameMap.CurrentLawyerKey, normalizedHeader);
		}

		[Fact]
		public void NormalizeMappingKeys_WhenInventoryAndCaseListHeaderUseCurrentKey_CanMatchWithoutAddingLegacyDefinitions()
		{
			IReadOnlyDictionary<string, CaseListFieldDefinition> fieldInventory =
				new Dictionary<string, CaseListFieldDefinition>(StringComparer.OrdinalIgnoreCase)
				{
					[FieldKeyRenameMap.CurrentLawyerKey] = new CaseListFieldDefinition
					{
						FieldKey = FieldKeyRenameMap.CurrentLawyerKey
					}
				};
			IReadOnlyDictionary<string, int> actualCaseListHeaders =
				new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
				{
					[FieldKeyRenameMap.CurrentLawyerKey] = 5
				};

			CaseListMappingDefinition mapping = Assert.Single(CaseListManagedKeyNormalizer.NormalizeMappingKeys(
				new[]
				{
					Mapping(FieldKeyRenameMap.LegacyLawyerKey, FieldKeyRenameMap.LegacyLawyerKey)
				}));

			Assert.True(fieldInventory.ContainsKey(mapping.SourceFieldKey));
			Assert.True(actualCaseListHeaders.ContainsKey(mapping.TargetHeaderName));
			Assert.False(fieldInventory.ContainsKey(FieldKeyRenameMap.LegacyLawyerKey));
			Assert.False(actualCaseListHeaders.ContainsKey(FieldKeyRenameMap.LegacyLawyerKey));
		}

		[Fact]
		public void NormalizeMappingKeys_WhenUnknownKeysExist_DoesNotMakeThemValid()
		{
			CaseListMappingDefinition mapping = Assert.Single(CaseListManagedKeyNormalizer.NormalizeMappingKeys(
				new[]
				{
					Mapping("未知キー", "未知ヘッダ")
				}));

			Assert.Equal("未知キー", mapping.SourceFieldKey);
			Assert.Equal("未知ヘッダ", mapping.TargetHeaderName);
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
