# TaskPaneRefreshOrchestrationService Phase 4 Safe Unit: R02

## 位置づけ

この文書は、TaskPane 表示回復領域の Phase 4 safe-first ownership separation の最初の実装記録です。

対象は R02 refresh precondition / fail-closed policy です。目的は `TaskPaneRefreshOrchestrationService` を削ることではなく、freeze line を守ったまま precondition decision の owner を少し外へ出すことです。

参照した正本:

- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/taskpane-display-recovery-current-state.md`
- `docs/taskpane-refresh-orchestration-responsibility-inventory.md`
- `docs/taskpane-refresh-orchestration-target-boundary-map.md`
- `docs/taskpane-display-recovery-freeze-line.md`

## GO / STOP 判断

GO と判断しました。

理由:

- R02 は refresh entry 直後、coordinator dispatch 前の precondition decision であり、retry / fallback / completion / foreground decision の順序に触れずに分離できます。
- `WorkbookOpen` window-dependent skip は既に `TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(...)` に固定済みです。
- protection gate の active window 判定と trace emit は `KernelHomeCasePaneSuppressionCoordinator` / `ICasePaneHostBridge` 側に残し、判定結果の扱いだけを policy decision に寄せられます。

## 分離したもの

- `TaskPaneRefreshPreconditionDecision` を追加しました。
- `TaskPaneRefreshPreconditionPolicy.DecideRefreshPrecondition(...)` を追加し、次の順序を policy 側の decision として固定しました。
  1. `WorkbookOpen` window-dependent skip。
  2. case foreground protection 中の `TaskPaneRefresh` ignore。
  3. proceed。
- `TaskPaneRefreshOrchestrationService` から nested `RefreshPreconditionEvaluator` と nested result object を削除しました。
- orchestration 側は policy decision を呼び、既存どおり skip action 名を trace / outcome normalization / WindowActivate downstream outcome に渡すだけにしました。

## 分離しなかったもの

- protection gate の active window 判定本体。
- `TaskPaneRefreshProtection action=ignore` trace の emit owner。
- `TaskPaneRefreshCoordinator` 内の defensive `WorkbookOpen` window-dependent skip。
- ready-show acceptance / display session / completion emit。
- ready-show retry、pending retry、active CASE fallback。
- foreground guarantee decision / outcome / trace。
- WindowActivate dispatch / downstream trace contract。
- workbook pane window resolve と `activateWorkbook` route policy。

## Freeze Line 確認

今回の変更では、次を変更していません。

- ready-show `attempt 1 -> 80ms retry attempt 2 -> pending retry fallback`。
- pending retry `400ms / 3 attempts` と workbook target / active CASE fallback。
- `WindowActivateDispatchOutcome.Dispatched != completion`。
- foreground outcome の status 意味。
- `case-display-completed` emit owner / one-time emit / completion 条件。
- display session boundary。
- protection / fail-closed 条件。
- trace 名と trace 意味。
- route contract。
- UI policy。
- COM restore 順序。
- Deploy 配下 / runtime Addins。

## Test 固定

`TaskPaneManagerOrchestrationPolicyTests` に R02 decision の boundary tests を追加しました。

- `WorkbookOpen` window-dependent skip が protection probe より先に判定されること。
- protection probe が block した場合は `ignore-during-protection` で fail-closed になること。
- どの gate も block しない場合は proceed になること。

## 次候補

次に安全そうな候補は R16 timer lifecycle boundary です。

理由:

- timer absent は no-op であり、owner 明確化の差分を小さくできます。
- ただし R06/R07/R08 retry sequence の順序や値は freeze line のため変更しません。

次点は R06 ready-show retry timer です。扱う場合は `80ms`、attempt 2 発火、cleanup owner を先に tests で固定します。
