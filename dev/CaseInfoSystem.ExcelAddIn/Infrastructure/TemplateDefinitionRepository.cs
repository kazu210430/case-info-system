using System.Collections.Generic;
using CaseInfoSystem.ExcelAddIn.Domain;
using CaseInfoSystem.ExcelAddIn.UI;

namespace CaseInfoSystem.ExcelAddIn.Infrastructure
{
    // クラス: Task Pane スナップショットを UI 用テンプレート定義へ変換する。
    // 責務: 特殊ボタンと文書ボタンを共通の TemplateDefinition 一覧へ詰め替える。
    internal sealed class TemplateDefinitionRepository
    {
        // メソッド: Task Pane スナップショットから TemplateDefinition 一覧を生成する。
        // 引数: snapshot - 変換元スナップショット。
        // 戻り値: TemplateDefinition 一覧。
        // 副作用: なし。
        internal IReadOnlyList<TemplateDefinition> Load(TaskPaneSnapshot snapshot)
        {
            var definitions = new List<TemplateDefinition>();
            if (snapshot == null)
            {
                return definitions;
            }

            foreach (TaskPaneActionDefinition action in snapshot.SpecialButtons)
            {
                definitions.Add(new TemplateDefinition
                {
                    Caption = action.Caption ?? string.Empty,
                    ActionKind = action.ActionKind ?? string.Empty,
                    Key = action.Key ?? string.Empty,
                    TabName = string.Empty,
                    RowIndex = 0,
                    BackColor = action.BackColor,
                    IsSpecialButton = true
                });
            }

            foreach (TaskPaneDocDefinition document in snapshot.DocButtons)
            {
                definitions.Add(new TemplateDefinition
                {
                    Caption = document.Caption ?? string.Empty,
                    ActionKind = document.ActionKind ?? string.Empty,
                    Key = document.Key ?? string.Empty,
                    TemplateFileName = document.TemplateFileName ?? string.Empty,
                    TabName = document.TabName ?? string.Empty,
                    RowIndex = document.RowIndex,
                    BackColor = document.BackColor,
                    IsSpecialButton = false
                });
            }

            return definitions;
        }
    }
}
