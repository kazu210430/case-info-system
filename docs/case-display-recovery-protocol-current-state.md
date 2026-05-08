# CASE Display Recovery Protocol Current State

## Implementation delta (2026-05-08 first safe unit)

- `TaskPaneRefreshOrchestrationService` now owns the created-case display session and is the only emitter of `case-display-completed`.
- `WorkbookTaskPaneReadyShowAttemptWorker` returns a normalized `WorkbookTaskPaneReadyShowAttemptOutcome`; it no longer completes the CASE display on the already-visible path.
- `TaskPaneRefreshCoordinator` returns a normalized `TaskPaneRefreshAttemptResult` with pane-visible / refresh-completed / foreground-terminal fields; it no longer emits `case-display-completed`.
- `TaskPaneHostFlowService` returns `TaskPaneHostFlowResult` so pane-visible source stays below the refresh coordinator without becoming display-completion ownership.
- `display-handoff-completed` is now emitted when the created-case display request is accepted by `TaskPaneRefreshOrchestrationService`; the old `display-handoff-open-completed` emit was removed from presentation/open-strategy code.
- Retry counts, ready-show timing, visibility recovery conditions, foreground recovery conditions, rebuild fallback, hidden session behavior, fail-closed conditions, and CASE creation behavior are unchanged.

## 位置づけ

この文書は、CASE display / recovery protocol の current-state 正本です。目的は、現行 `main` で実際に動いている owner と順序を、観測ログと現コードを根拠に固定することです。

- 基準コード: 現行 `main` / `origin/main` 一致時点の `16e3a415a929b60eee45e13c606152d04d83676b`
- 構成全体の前提: `docs/architecture.md`
- 主要フローの前提: `docs/flows.md`
- UI 制御の前提: `docs/ui-policy.md`
- review-first safe unit の前提: `docs/codex-review-first-safe-unit-plan.md`
- ready-show / recovery 観測整理: `docs/readyshow-recovery-observation-points-2026-05-08.md`

この文書は理想設計を書き換えるものではありません。いまの protocol がどの service に分散し、どこで重複し、どこが未正本化のまま動いているかを current-state として残すための文書です。protocol redesign、owner 統合、retry / recovery 条件変更はこの文書の対象外です。

## この文書で固定すること

この文書では、次だけを固定します。

1. CASE workbook open から CASE display completed までの current-state の順序。
2. 各 protocol unit を、現実には誰が ownership しているか。
3. owner が 1 箇所に閉じていない箇所、重複観測される箇所、暗黙 protocol になっている箇所。
4. 次に redesign に入るなら、どこを最初の安全単位として切るべきか。

この文書では、次は固定しません。

- 表示不安定の根本原因
- 正しいあるべき設計
- retry 値や recovery 条件の変更根拠
- hidden session 修正
- CASE 作成本体の修正

## 用語の current-state

現行コードでは、似た言葉が別概念として動いています。この文書では次の意味で使います。

- `display handoff`
  - CASE workbook を hidden reopen し、表示系 protocol へ渡す境界です。
  - first safe unit 後は `TaskPaneRefreshOrchestrationService` が ready-show acceptance 時に `display-handoff-completed` を記録します。
- `ready-show enqueue`
  - 「ready になったら CASE pane を見せる」要求を retry 可能な ready-show queue に入れる段階です。
- `ready-show attempt`
  - ready-show queue から 1 回分の表示試行を実行する段階です。
- `pane visible`
  - CASE pane host が reuse または refresh 後に実際に shown された状態です。
  - current-state では `TaskPaneHostFlowService` が `taskpane-reused-shown` / `taskpane-refreshed-shown` を出します。
- `foreground guarantee`
  - refresh 後に Excel / workbook window を最終的に前面へ戻す protocol です。
  - current-state では decision owner と execution owner が分かれています。
- `CASE display completed`
  - first safe unit 後は `TaskPaneRefreshOrchestrationService` が created-case display session を閉じる唯一の owner です。
  - `WorkbookTaskPaneReadyShowAttemptWorker` と `TaskPaneRefreshCoordinator` は normalized outcome を返し、`case-display-completed` を emit しません。

重要なのは、`CASE display completed`、`pane visible`、`foreground guarantee` は同じ意味ではないことです。現行 code では別 service、別タイミング、別条件で扱われています。

## Current-State Flow

### 1. CASE workbook open

- `KernelCasePresentationService.OpenCreatedCase(...)` が `case-workbook-open-started` を記録します。
- `KernelCasePresentationService.OpenCreatedCaseWorkbook(...)` は作成種別に応じて workbook open strategy を選びます。
- `NewCaseDefault` と `CreateCaseSingle` では `CaseWorkbookOpenStrategy.OpenHiddenForCaseDisplay(...)` が選ばれます。
- `CaseWorkbookOpenStrategy.OpenHiddenForCaseDisplay(...)` は shared display state を適用し、opened workbook window を hidden にし、必要に応じて前の window を復元します。
- open 完了後、`KernelCasePresentationService` は `case-workbook-open-completed` を記録します。

### 2. display handoff

- hidden reopen 後、`KernelCasePresentationService` が ready-show request を発行します。
- `TaskPaneRefreshOrchestrationService.ShowWorkbookTaskPaneWhenReady(...)` は request を受理し、created-case display session を開始します。
- この acceptance 境界で `created-case-display-session-started` と `display-handoff-completed` を記録します。

つまり first safe unit 後は、display handoff completion の観測 owner は `TaskPaneRefreshOrchestrationService` に寄っています。

### 3. initial recovery and ready-show request

- `KernelCasePresentationService.ShowCreatedCase(...)` は ready-show 前の初期 recovery を調停します。
- ここでは `WorkbookWindowVisibilityService.EnsureVisible(...)` による workbook window 可視化評価と、`ExcelWindowRecoveryService.TryRecoverWorkbookWindowWithoutShowing(..., bringToFront: false)` による initial recovery が実行されます。
- その後 `initial-recovery-completed` が記録されます。
- `KernelCasePresentationService.ExecuteDeferredPresentationEnhancements(...)` は suppression release、post-release suppression prepare、ready-show request を順番に実行します。
- ここで `post-release-suppression-prepared` と `ready-show-requested` が記録されます。

### 4. ready-show enqueue

- `TaskPaneRefreshOrchestrationService.ShowWorkbookTaskPaneWhenReady(...)` が ready-show queue の入口です。
- ここで `ready-show-enqueued` が記録され、`WorkbookTaskPaneReadyShowAttemptWorker.ShowWhenReady(...)` へ渡されます。
- 同じ orchestration service には、workbook-open 直後など window 依存 refresh が成立しない場合の fallback handoff、timer fallback、ready retry 調停も残っています。

### 5. ready-show attempt

- `WorkbookTaskPaneReadyShowAttemptWorker.TryShowWorkbookTaskPaneOnce(...)` が 1 attempt の owner です。
- ここで `ready-show-attempt` と `ready-show-attempt-result` が記録されます。
- attempt 前には `EnsureWorkbookWindowVisibleForTaskPaneDisplay(...)` が走り、初回は `WorkbookWindowVisibilityService.EnsureVisible(...)` を通ります。
- ready 条件が不足していれば orchestration service 側の retry schedule へ戻り、`TaskPane wait-ready retry scheduled` / `TaskPane wait-ready retry firing` が plain log として残ります。

### 6. taskpane already visible path

- `WorkbookTaskPaneReadyShowAttemptWorker` は、既存 CASE pane が同じ workbook window に対して visible かどうかを判定します。
- visible であれば `taskpane-already-visible` を記録し、refresh せずに success として抜けます。
- first safe unit 後は `WorkbookTaskPaneReadyShowAttemptWorker` が `WorkbookTaskPaneReadyShowAttemptOutcome` を返し、`case-display-completed` は orchestration 側で成立します。

この path では、pane already-visible の検知は worker 側に残しつつ、CASE display completed 判定は `TaskPaneRefreshOrchestrationService` 側へ移しています。

### 7. taskpane refresh path

- already-visible でなければ、`WorkbookTaskPaneReadyShowAttemptWorker` は refresh delegate を呼びます。
- 入口 owner は `TaskPaneRefreshCoordinator.TryRefreshTaskPane(...)` です。
- coordinator は `TaskPaneRefreshPreconditionPolicy` に従って `WorkbookOpen` 直後の window-dependent refresh を skip する shared boundary を持ちます。
- coordinator は context 解決前に必要なら `ExcelWindowRecoveryService` で workbook window recovery を行い、pane refresh 用 window と context を解決します。
- refresh 本体では `taskpane-refresh-started` と `taskpane-refresh-completed` が記録されます。

### 8. rebuild fallback と snapshot source decision

- rebuild fallback は ready-show の前段ではなく、taskpane refresh path の内部で発生します。
- 呼び出し順は次です。
  - `TaskPaneRefreshCoordinator.TryRefreshTaskPane(...)`
  - `TaskPaneManager.RefreshPane(...)`
  - `TaskPaneHostFlowService.RenderAndShowHostForRefresh(...)`
  - `TaskPaneManager.RenderCaseHost(...)`
  - `CasePaneSnapshotRenderService.Render(...)`
  - `TaskPaneSnapshotBuilderService.BuildSnapshotText(...)`
- snapshot source decision の current-state owner は `TaskPaneSnapshotBuilderService` です。
- build order は `CaseCache -> BaseCache -> BaseCacheFallback -> MasterListRebuild` です。
- `Task pane snapshot rebuild fallback selected...` と `Task pane snapshot MasterListRebuild started...` が rebuild fallback の protocol 上の観測点です。

したがって rebuild fallback は refresh path 内の render/snapshot protocol に属し、ready-show enqueue や foreground guarantee の owner ではありません。

### 9. pane visible

- actual pane visible transition は `TaskPaneHostFlowService` が持ちます。
- host reuse path では `taskpane-reused-shown` を記録します。
- refresh render path では `taskpane-refreshed-shown` を記録します。

これは `case-display-completed` と同義ではありません。first safe unit 後も pane visible trace は host-flow 層に残し、display completion trace は orchestration 層に集約しています。

### 10. foreground guarantee

- refresh 後の foreground guarantee decision owner は `TaskPaneRefreshCoordinator` です。
- coordinator は `foreground-recovery-decision` を記録し、必要時のみ `GuaranteeFinalForegroundAfterRefresh(...)` に進みます。
- final guarantee の execution owner は `ExcelWindowRecoveryService` です。
- coordinator は `final-foreground-guarantee-started` / `final-foreground-guarantee-completed` を記録し、その内部で `ExcelWindowRecoveryService.TryRecoverWorkbookWindow(..., bringToFront: true)` または active version が実行されます。

つまり foreground guarantee は 1 owner ではなく、decision と execution が別 service に分かれています。

### 11. visibility recovery

- visibility recovery も 1 owner ではありません。
- `WorkbookWindowVisibilityService`
  - workbook window が visible かを評価し、必要なら `window.Visible = true` を行います。
  - ready-show 前と初期 recovery 前の workbook-window visibility owner です。
- `ExcelWindowRecoveryService`
  - workbook window 再解決、window 可視化、minimized restore、application visible ensure、activation、foreground promotion まで持ちます。
  - `Excel window recovery evaluated...` と `Excel window recovery mutation trace...` が execution trace です。

current-state の visibility recovery は、lightweight workbook visibility ensure と、full Excel window recovery primitive の 2 層に跨っています。

### 12. CASE display completed

- first safe unit 後は、`case-display-completed` の定義 owner は `TaskPaneRefreshOrchestrationService` です。
- `WorkbookTaskPaneReadyShowAttemptWorker`
  - already-visible success を `WorkbookTaskPaneReadyShowAttemptOutcome` として返します。
- `TaskPaneRefreshCoordinator`
  - refresh / foreground terminal を `TaskPaneRefreshAttemptResult` として返します。

このため、CASE display completed は「pane visible になった瞬間」や「refresh completed の別名」ではなく、同一 created-case display session の pane visible と foreground terminal を orchestration 側が確認して成立させます。

## Protocol Unit と Current-State Owner

| Protocol unit | Current-state owner | 補足 |
| --- | --- | --- |
| CASE workbook open | `KernelCasePresentationService` と `CaseWorkbookOpenStrategy` | presentation が flow 入口、hidden reopen 実処理は strategy |
| WorkbookOpen | `WorkbookLifecycleCoordinator` | event capture と lifecycle / refresh dispatch owner |
| WorkbookActivate | `WorkbookLifecycleCoordinator` | suppression / protection を見た上で refresh dispatch |
| WindowActivate event capture | `ThisAddIn.HandleWindowActivateEvent(...)` | Excel event capture と trace |
| WindowActivate refresh dispatch | `WindowActivatePaneHandlingService` | request 化して refresh へ渡す |
| display handoff | `TaskPaneRefreshOrchestrationService` | `display-handoff-completed` は ready-show acceptance 側 |
| ready-show request | `KernelCasePresentationService` | `ready-show-requested` |
| ready-show enqueue | `TaskPaneRefreshOrchestrationService` | `ready-show-enqueued` |
| ready-show attempt | `WorkbookTaskPaneReadyShowAttemptWorker` | retry 可能な 1 attempt owner |
| taskpane already visible path | `WorkbookTaskPaneReadyShowAttemptWorker` | early-complete path |
| taskpane refresh path | `TaskPaneRefreshCoordinator` | context resolve、refresh、foreground decision |
| pane visible | `TaskPaneHostFlowService` | `taskpane-reused-shown` / `taskpane-refreshed-shown` |
| foreground guarantee decision | `TaskPaneRefreshCoordinator` | refresh 後に走るかどうかの判断 |
| foreground guarantee execution | `ExcelWindowRecoveryService` | app/window/foreground recovery primitive |
| visibility recovery | `WorkbookWindowVisibilityService` と `ExcelWindowRecoveryService` | workbook visibility と full recovery に分裂 |
| rebuild fallback | `TaskPaneSnapshotBuilderService` | refresh path 内の render/snapshot protocol |
| snapshot source decision | `TaskPaneSnapshotBuilderService` | `CaseCache -> BaseCache -> BaseCacheFallback -> MasterListRebuild` |
| CASE display completed | `TaskPaneRefreshOrchestrationService` | created-case display session の terminal owner |

## ownership 混在・重複・暗黙 protocol

現行 code と観測から、次を current-state の混在として固定します。

- `display-handoff-open-completed` の重複 emit は first safe unit で削除済み。
  - `display-handoff-completed` は `TaskPaneRefreshOrchestrationService` が ready-show acceptance 側で emit する。
- `WindowActivate` owner が event capture と refresh dispatch に分裂している。
  - capture は `ThisAddIn.HandleWindowActivateEvent(...)`
  - dispatch は `WindowActivatePaneHandlingService`
- ready-show 系 owner が 3 層に分裂している。
  - request は `KernelCasePresentationService`
  - enqueue / fallback / retry 調停は `TaskPaneRefreshOrchestrationService`
  - attempt / early-complete は `WorkbookTaskPaneReadyShowAttemptWorker`
- `CASE display completed` owner は `TaskPaneRefreshOrchestrationService` に集約済み。
  - already-visible completion 材料は worker outcome
  - refresh / foreground 材料は coordinator outcome
  - actual pane visible trace は host-flow service
- foreground guarantee が decision と execution に分裂している。
  - decision は `TaskPaneRefreshCoordinator`
  - execution は `ExcelWindowRecoveryService`
- visibility recovery が複数 service に跨っている。
  - workbook visible ensure は `WorkbookWindowVisibilityService`
  - application/window/foreground recovery は `ExcelWindowRecoveryService`
- rebuild fallback は display protocol の前段ではなく、refresh path の深い内部に埋まっている。
  - current-state では ready-show から直接は見えず、render/snapshot path に入って初めて発火する。
- refresh source owner が未正本化である。
  - `TaskPaneRefreshCoordinator` の `taskpane-refresh-started` / `completed` では `refreshSource` が実質 `reason` の再掲になっている。
  - source を誰が決め、reason とどう区別するかは current-state で 1 owner に閉じていない。

## 最重要 owner 不明・未正本化箇所

current-state の中でも、次が protocol redesign 前に最も重要な「owner 不明」または「未正本化」箇所です。

### 1. CASE display completed definition owner

- first safe unit 後は `TaskPaneRefreshOrchestrationService` の 1 箇所で完了扱いになります。
- pane visible trace と foreground guarantee trace は別 service にあるため、「何をもって CASE 表示完了とするか」の正本 owner がありません。

### 2. refresh source owner

- `reason` は caller 起点で流れてきますが、`refreshSource` は protocol 上の正規化された source owner を持っていません。
- snapshot source、refresh trigger source、display completion source が別概念のまま同列に見えやすく、current-state docs でも未正本化として扱うべき領域です。

### 3. display handoff completion owner

- hidden reopen handoff を完了させる owner と、その completion trace owner が一致していません。
- 同名 trace の重複により、観測上は「どちらが protocol owner か」が曖昧です。

## 次に protocol redesign へ入るなら最初の安全単位候補

最初の安全単位候補は、`CASE display completed` の定義 owner を current-state 上で 1 箇所に寄せる準備です。

理由は次です。

- 表示完了、pane visible、foreground guarantee が別概念のまま分離している。
- completion trace は `TaskPaneRefreshOrchestrationService` に集約済みです。worker / coordinator / host-flow は outcome と visible trace を返す側に残します。
- recovery 条件、retry 条件、visibility 制御、ready-show timing を変えずに、まず completion definition の ownership だけを切り出して観測・文書化しやすい。

この候補は redesign 実施を意味しません。current-state としては、「最初に触るなら completion definition boundary を安全単位化するのが最も説明可能」という提案に留めます。

## 次に target-state 設計で決めるべき論点

この節は target-state を確定するものではありません。current-state で見えている分裂状態を前提に、次に設計判断が必要な論点だけを列挙します。

### 1. CASE display completed definition

- current-state の分裂状態
  - `WorkbookTaskPaneReadyShowAttemptWorker` が already-visible path で completion を成立させる。
  - `TaskPaneRefreshCoordinator` が refresh path で completion を成立させる。
  - `TaskPaneHostFlowService` は pane visible を記録するが、completion owner ではない。
  - `TaskPaneRefreshCoordinator` と `ExcelWindowRecoveryService` の foreground guarantee 完了とも一致していない。
- 次に設計判断が必要な点
  - completion を 1 owner に寄せるのか。
  - completion 判定を `pane visible`、`refresh completed`、`foreground guarantee completed` のどこに結び付けるのか。
  - already-visible path と refresh path の completion semantics を同一にするのか。

### 2. ready-show orchestration owner

- current-state の分裂状態
  - request は `KernelCasePresentationService`
  - enqueue / fallback / retry 調停は `TaskPaneRefreshOrchestrationService`
  - attempt / early-complete は `WorkbookTaskPaneReadyShowAttemptWorker`
- 次に設計判断が必要な点
  - ready-show protocol の owner を request から completion まで 1 本の orchestration として扱うのか。
  - fallback / retry / attempt を同一 owner に閉じるのか、それとも queue owner と attempt owner を分けたまま明示するのか。
  - ready-show が CASE display protocol のどこからどこまでを責務に含むのか。

### 3. foreground guarantee owner

- current-state の分裂状態
  - decision は `TaskPaneRefreshCoordinator`
  - execution は `ExcelWindowRecoveryService`
  - completion trace は coordinator 側にあり、mutation trace は recovery service 側にある。
- 次に設計判断が必要な点
  - foreground guarantee を 1 protocol unit として誰が ownership するのか。
  - decision と execution を分けたままにするなら、どちらを正本 owner と呼ぶのか。
  - CASE display completed との前後関係を protocol 上でどう固定するのか。

### 4. visibility recovery owner

- current-state の分裂状態
  - `WorkbookWindowVisibilityService` が workbook window visible ensure を持つ。
  - `ExcelWindowRecoveryService` が window 再解決、application visible ensure、restore、activation、foreground promotionを持つ。
  - `KernelCasePresentationService` と `WorkbookTaskPaneReadyShowAttemptWorker` がそれらを別タイミングで呼ぶ。
- 次に設計判断が必要な点
  - visibility recovery を 1 protocol としてまとめるのか、lightweight ensure と full recovery を別 protocol とするのか。
  - visible ensure と foreground promotion の境界をどこで切るのか。
  - caller 側が recovery primitive を直接組み合わせる current-state を維持するのか。

### 5. rebuild fallback owner

- current-state の分裂状態
  - rebuild fallback decision 自体は `TaskPaneSnapshotBuilderService`
  - ただし protocol 上は `TaskPaneRefreshCoordinator -> TaskPaneManager -> TaskPaneHostFlowService -> CasePaneSnapshotRenderService` の深い内部で発火する。
  - ready-show / refresh orchestration 側からは rebuild fallback が直接見えにくい。
- 次に設計判断が必要な点
  - rebuild fallback を display protocol の明示 unit として上位へ持ち上げるのか。
  - snapshot build owner と display protocol owner の境界をどこに置くのか。
  - rebuild fallback 開始時点を protocol trace 上どこで定義するのか。

### 6. refresh source owner

- current-state の分裂状態
  - refresh trigger の `reason` は caller 起点で流れる。
  - `TaskPaneRefreshCoordinator` の `refreshSource` は実質 `reason` の再掲で、source 正規化 owner がない。
  - snapshot source は `TaskPaneSnapshotBuilderService` が別に決めている。
- 次に設計判断が必要な点
  - refresh trigger source と snapshot source を同じ語で扱うのか、別概念として分けるのか。
  - refresh source の正本 owner を coordinator に置くのか、もっと上流に置くのか。
  - 観測上の `reason` と protocol 上の `source` をどう切り分けるのか。

### 7. WindowActivate 系 owner

- current-state の分裂状態
  - Excel event capture と trace は `ThisAddIn.HandleWindowActivateEvent(...)`
  - request 化と refresh dispatch は `WindowActivatePaneHandlingService`
  - `WorkbookActivate` とは別に `WorkbookLifecycleCoordinator` も refresh dispatch を持つ。
- 次に設計判断が必要な点
  - `WindowActivate` を 1 protocol unit として誰が owner になるのか。
  - event capture owner と refresh dispatch owner を分け続けるのか。
  - `WorkbookActivate` と `WindowActivate` の責務境界を protocol 上どこで分離するのか。

## 今回の current-state に含めないこと

- 表示不安定の原因断定
- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` の timing 変更
- ready-show retry 値の調整
- foreground recovery 条件の変更
- visibility 制御の変更
- hidden session 修正
- CASE 作成本体の修正

## 一言まとめ

現行の CASE display / recovery protocol は、単一 owner の直線的な flow ではありません。display handoff、ready-show、refresh、foreground guarantee、visibility recovery、rebuild fallback、display completion が複数 service に分散し、いくつかは重複観測と未正本化のまま連結しています。

この文書では、その分散を問題視する前に、まず current-state として固定しました。次に進むなら、completion definition から安全単位で ownership を整理するのが最初の候補です。
