using System;
using System.Collections.Generic;
using System.Drawing;

namespace CaseInfoSystem.ExcelAddIn.UI
{
    // クラス: Task Pane 全体のスナップショットを表す。
    // 責務: 特殊ボタン、タブ、文書ボタンと表示補助情報をまとめて保持する。
    internal sealed class TaskPaneSnapshot
    {
        // メソッド: Task Pane スナップショットを初期化する。
        // 引数: なし。
        // 戻り値: なし。
        // 副作用: 各コレクションを生成する。
        internal TaskPaneSnapshot()
        {
            SpecialButtons = new List<TaskPaneActionDefinition>();
            Tabs = new List<TaskPaneTabDefinition>();
            DocButtons = new List<TaskPaneDocDefinition>();
        }

        internal string WorkbookName { get; set; }
        internal string WorkbookPath { get; set; }
        internal bool HasError { get; set; }
        internal string ErrorMessage { get; set; }
        internal int PreferredPaneWidth { get; set; }
        internal List<TaskPaneActionDefinition> SpecialButtons { get; }
        internal List<TaskPaneTabDefinition> Tabs { get; }
        internal List<TaskPaneDocDefinition> DocButtons { get; }
    }

    // クラス: Task Pane ボタン共通の表示定義を表す。
    // 責務: 特殊ボタン・文書ボタン共通の表示位置や配色を保持する。
    internal class TaskPaneActionDefinition
    {
        internal string Name { get; set; }
        internal string Caption { get; set; }
        internal string ActionKind { get; set; }
        internal string Key { get; set; }
        internal Color BackColor { get; set; }
        internal int Left { get; set; }
        internal int Top { get; set; }
        internal int Width { get; set; }
        internal int Height { get; set; }
    }

    // クラス: Task Pane のタブ定義を表す。
    // 責務: タブ順、タブ名、背景色を保持する。
    internal sealed class TaskPaneTabDefinition
    {
        internal int Order { get; set; }
        internal string TabName { get; set; }
        internal Color BackColor { get; set; }
    }

    // クラス: Task Pane 上の文書ボタン定義を表す。
    // 責務: 共通ボタン情報に加えて雛形名とタブ情報を保持する。
    internal sealed class TaskPaneDocDefinition : TaskPaneActionDefinition
    {
        internal string TemplateFileName { get; set; }

        internal string TabName { get; set; }
        internal int RowIndex { get; set; }

        internal Color FillColor
        {
            get { return BackColor; }
            set { BackColor = value; }
        }
    }

    // クラス: Task Pane 上で押下されたアクションを通知するイベント引数。
    // 責務: 動作種別と対象キーを保持する。
    internal sealed class TaskPaneActionEventArgs : EventArgs
    {
        // メソッド: イベント引数を初期化する。
        // 引数: actionKind - 動作種別, key - 対象キー。
        // 戻り値: なし。
        // 副作用: 内部状態を初期化する。
        internal TaskPaneActionEventArgs(string actionKind, string key)
        {
            ActionKind = actionKind ?? string.Empty;
            Key = key ?? string.Empty;
        }

        internal string ActionKind { get; }
        internal string Key { get; }
    }
}
