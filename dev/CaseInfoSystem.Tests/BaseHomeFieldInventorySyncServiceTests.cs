using System.Linq;
using CaseInfoSystem.ExcelAddIn.App;
using Xunit;

namespace CaseInfoSystem.Tests
{
	public sealed class BaseHomeFieldInventorySyncServiceTests
	{
		[Fact]
		public void BuildPlan_WhenBaseKeyChanged_UpdatesOnlyProposedFieldKey()
		{
			BaseHomeFieldInventorySyncService.SyncPlan plan = BaseHomeFieldInventorySyncService.BuildPlan(
				new[]
				{
					BaseRow(1, "新キー")
				},
				new[]
				{
					InventoryRow(2, "B1", "旧キー")
				},
				new string[0]);

			Assert.True(plan.CanApply);
			BaseHomeFieldInventorySyncService.FieldInventoryUpdate update = Assert.Single(plan.Updates);
			Assert.Equal(1, update.BaseHomeRowNumber);
			Assert.Equal(2, update.FieldInventoryRowNumber);
			Assert.Equal("旧キー", update.OldFieldKey);
			Assert.Equal("新キー", update.NewFieldKey);
			Assert.Equal(0, plan.UnchangedCount);
		}

		[Fact]
		public void BuildPlan_WhenKeyUnchanged_DoesNotCreateUpdate()
		{
			BaseHomeFieldInventorySyncService.SyncPlan plan = BaseHomeFieldInventorySyncService.BuildPlan(
				new[]
				{
					BaseRow(1, "顧客_住所")
				},
				new[]
				{
					InventoryRow(2, "B1", "顧客_住所")
				},
				new string[0]);

			Assert.True(plan.CanApply);
			Assert.Empty(plan.Updates);
			Assert.Equal(1, plan.UnchangedCount);
		}

		[Fact]
		public void BuildPlan_WhenBaseKeyIsBlank_FailsClosed()
		{
			BaseHomeFieldInventorySyncService.SyncPlan plan = BaseHomeFieldInventorySyncService.BuildPlan(
				new[]
				{
					BaseRow(1, "")
				},
				new[]
				{
					InventoryRow(2, "B1", "顧客_名前")
				},
				new string[0]);

			Assert.False(plan.CanApply);
			Assert.Contains(plan.Errors, error => error.Contains("空欄"));
			Assert.Empty(plan.Updates);
		}

		[Fact]
		public void BuildPlan_WhenBaseKeyIsDuplicated_FailsClosed()
		{
			BaseHomeFieldInventorySyncService.SyncPlan plan = BaseHomeFieldInventorySyncService.BuildPlan(
				new[]
				{
					BaseRow(1, "重複キー"),
					BaseRow(2, "重複キー")
				},
				new[]
				{
					InventoryRow(2, "B1", "旧キー1"),
					InventoryRow(3, "B2", "旧キー2")
				},
				new string[0]);

			Assert.False(plan.CanApply);
			Assert.Contains(plan.Errors, error => error.Contains("重複"));
			Assert.Single(plan.Updates);
		}

		[Fact]
		public void BuildPlan_WhenImportantKeyWouldChange_FailsClosed()
		{
			BaseHomeFieldInventorySyncService.SyncPlan plan = BaseHomeFieldInventorySyncService.BuildPlan(
				new[]
				{
					BaseRow(1, "顧客_名前_変更")
				},
				new[]
				{
					InventoryRow(2, "B1", "顧客_名前")
				},
				new[] { "顧客_名前" });

			Assert.False(plan.CanApply);
			Assert.Contains(plan.Errors, error => error.Contains("重要キー"));
			Assert.Empty(plan.Updates);
		}

		[Fact]
		public void BuildPlan_WhenFinalFieldKeyWouldDuplicate_FailsClosed()
		{
			BaseHomeFieldInventorySyncService.SyncPlan plan = BaseHomeFieldInventorySyncService.BuildPlan(
				new[]
				{
					BaseRow(1, "既存キー")
				},
				new[]
				{
					InventoryRow(2, "B1", "旧キー"),
					InventoryRow(3, "B99", "既存キー")
				},
				new string[0]);

			Assert.False(plan.CanApply);
			Assert.Contains(plan.Errors, error => error.Contains("同期後"));
		}

		[Fact]
		public void BuildPlan_WhenInventoryRowIsOutsideBaseRange_WarnsButKeepsExistingRow()
		{
			BaseHomeFieldInventorySyncService.SyncPlan plan = BaseHomeFieldInventorySyncService.BuildPlan(
				new[]
				{
					BaseRow(1, "顧客_名前")
				},
				new[]
				{
					InventoryRow(2, "B1", "顧客_名前"),
					InventoryRow(3, "B2", "追加キー")
				},
				new string[0]);

			Assert.True(plan.CanApply);
			Assert.Empty(plan.Updates);
			Assert.Single(plan.Warnings);
			Assert.Contains("既存行は削除しません", plan.Warnings.Single());
		}

		private static BaseHomeFieldInventorySyncService.BaseHomeFieldKeyRow BaseRow(int rowNumber, string fieldKey)
		{
			return new BaseHomeFieldInventorySyncService.BaseHomeFieldKeyRow
			{
				RowNumber = rowNumber,
				FieldKey = fieldKey
			};
		}

		private static BaseHomeFieldInventorySyncService.FieldInventoryRow InventoryRow(int rowNumber, string sourceCell, string proposedFieldKey)
		{
			return new BaseHomeFieldInventorySyncService.FieldInventoryRow
			{
				RowNumber = rowNumber,
				SourceCell = sourceCell,
				ProposedFieldKey = proposedFieldKey
			};
		}
	}
}
