# TaskPane Refactor Current State

## 目的と前提

この文書は、優先度Aの TaskPane 側リファクタについて、現時点の到達点と積み残しをゼロから棚卸しし直さず再開できるように固定するための作業基準です。

- この文書は、`main` に `83563bc1cd90cdbafcfa5e6579f4bcd41c76389c` を含む状態を基準に整理する
- 設計全体の正本は `docs/taskpane-architecture.md`
- 構成全体の前提は `docs/architecture.md`
- 対象フローの前提は `docs/flows.md`
- UI 制御方針の前提は `docs/ui-policy.md`

この文書では、新しい設計提案は増やさず、確認できた到達点と保留事項だけを記録します。

## 現在の到達点

TaskPane 側で分離済みの責務は次のとおりです。

### `TaskPaneActionDispatcher`

- CASE pane の業務アクション起動を担う
- post-action refresh 判定との接続を担う

### `TaskPaneHostRegistry`

- host lifecycle を担う
- `register / replace / remove / dispose` を担う

### `TaskPaneRefreshFlowCoordinator`

- `RefreshPane` 主経路を担う
- `precondition -> host解決 -> CASE host再利用 -> render/show` の順序を担う

### `TaskPaneDisplayCoordinator`

- 表示判断の正本を担う
- display request 判定を担う
- visible pane 判定を担う
- role / workbook / window に基づく表示判断を担う

補足:

- `TaskPaneDisplayCoordinator` は別ファイル化済み
- `TaskPaneHostRegistry` / `TaskPaneActionDispatcher` / `TaskPaneRefreshFlowCoordinator` は、責務分離済みだが現時点では `TaskPaneManager` 内の nested class として維持している

## `TaskPaneManager` に残す責務

`TaskPaneManager` は、現時点では次を担う状態として整理する。

- TaskPane 全体の入口
- refresh / display / host / action 各 coordinator への接続
- 既存 VSTO / UI 境界との接続
- render 入口の保持

補足:

- `ThisAddIn.RequestTaskPaneDisplayForTargetWindow(...)` から入る表示経路は維持する
- role 別 render 入口は `TaskPaneManager` に残し、境界の安定を優先する

## 意図的に残したもの

以下は、現時点では意図的に残す。

### `ThisAddIn.cs`

- VSTOイベント境界なので最後に薄くする
- 今の段階で触るとイベントタイミング差分が出やすい

### TaskPane 関連クラスの別ファイル分割

- すでに正本があるものは正本へ寄せる
- nested class は、安定後に必要なら別ファイル化する

### Kernel 側の残作業

- TaskPane 側の固定後、同じ粒度で進める

## 今後の作業方針

- ゼロから棚卸しをやり直さない
- 既存の到達点を前提に次の最小単位を進める
- 1回に1責務だけ切る
- build / test / 実機確認 / main固定を必ず挟む
- `.md` は作業再開時の基準点として使う

## 禁止事項

- `TaskPaneManager` に display coordinator を二重作成しない
- `TaskPaneDisplayCoordinator.cs` が表示判断の正本
- `TaskPaneHostRegistry` に表示判断を混ぜない
- `TaskPaneActionDispatcher` に表示判断を混ぜない
- `TaskPaneRefreshFlowCoordinator` に host lifecycle や action 起動を混ぜない
- UI表示タイミングを変えない
- `RequestTaskPaneDisplayForTargetWindow(...)` の呼び出し位置を変えない

## 未確認として扱う事項

- retry 秒数や protection 秒数の正式な仕様根拠は未確認
- nested class の別ファイル化時期は未確定
- `ThisAddIn.cs` をどの順で薄くするかの最終順序は未確定

## 参照元

- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/taskpane-architecture.md`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneManager.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneDisplayCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/ThisAddIn.cs`
