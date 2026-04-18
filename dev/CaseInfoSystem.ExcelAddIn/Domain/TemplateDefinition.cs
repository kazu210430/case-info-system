using System.Drawing;

namespace CaseInfoSystem.ExcelAddIn.Domain
{
    // クラス: Task Pane 上の1ボタン分の表示定義を表す。
    // 責務: 文書ボタン・特殊ボタン共通の表示情報と動作情報を保持する。
    internal sealed class TemplateDefinition
    {
        internal string Caption { get; set; }
        internal string ActionKind { get; set; }
        internal string Key { get; set; }
        internal string TemplateFileName { get; set; }
        internal string TabName { get; set; }
        internal int RowIndex { get; set; }
        internal Color BackColor { get; set; }
        internal bool IsSpecialButton { get; set; }
    }
}
