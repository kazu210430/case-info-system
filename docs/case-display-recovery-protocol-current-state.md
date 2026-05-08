# CASE Display Recovery Protocol Current State

## Implementation delta (2026-05-08 first safe unit)

- `TaskPaneRefreshOrchestrationService` now owns the created-case display session and is the only emitter of `case-display-completed`.
- `WorkbookTaskPaneReadyShowAttemptWorker` returns a normalized `WorkbookTaskPaneReadyShowAttemptOutcome`; it no longer completes the CASE display on the already-visible path.
- `TaskPaneRefreshCoordinator` returns a normalized `TaskPaneRefreshAttemptResult` with pane-visible / refresh-completed / foreground-terminal fields; it no longer emits `case-display-completed`.
- `TaskPaneHostFlowService` returns `TaskPaneHostFlowResult` so pane-visible source stays below the refresh coordinator without becoming display-completion ownership.
- `display-handoff-completed` is now emitted when the created-case display request is accepted by `TaskPaneRefreshOrchestrationService`; the old `display-handoff-open-completed` emit was removed from presentation/open-strategy code.
- Retry counts, ready-show timing, visibility recovery conditions, foreground recovery conditions, rebuild fallback, hidden session behavior, fail-closed conditions, and CASE creation behavior are unchanged.
- Completion status: merged to `main` at `e41feb5d607f79077e112a1945e81ac0a76d95a4` and verified on the actual Excel runtime.

## Implementation delta (2026-05-08 foreground guarantee first safe unit)

- `TaskPaneRefreshOrchestrationService` now owns foreground guarantee decision, normalized outcome, and `foreground-recovery-decision` / `final-foreground-guarantee-*` trace emission.
- `TaskPaneRefreshCoordinator` returns refresh raw facts and exposes the existing foreground recovery execution bridge; it no longer normalizes foreground outcome or emits final foreground guarantee traces.
- `ForegroundGuaranteeOutcome` now records `NotRequired` / `SkippedAlreadyVisible` / `SkippedNoKnownTarget` / `RequiredSucceeded` / `RequiredDegraded` / `RequiredFailed` / `Unknown`.
- `case-display-completed` remains owned by `TaskPaneRefreshOrchestrationService` and now consumes `ForegroundGuaranteeOutcome.IsTerminal` plus `IsDisplayCompletable`.
- Retry counts, ready-show timing, visibility recovery conditions, foreground recovery execution conditions, rebuild fallback, `WindowActivate` behavior, hidden session behavior, and CASE creation behavior are unchanged.

## 位置づけ

この文書は、CASE display / recovery protocol の current-state 正本です。目的は、現行 `main` で実際に動いている owner と順序を、観測ログと現コードを根拠に固定することです。

- 基準コード: 現行 `main` / `origin/main` 一致時点の `e41feb5d607f79077e112a1945e81ac0a76d95a4`
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

- refresh 後の foreground guarantee decision / outcome owner は `TaskPaneRefreshOrchestrationService` です。
- orchestration は `foreground-recovery-decision` を記録し、必要時のみ coordinator の execution bridge から `ExcelWindowRecoveryService` へ進みます。
- final guarantee の execution owner は `ExcelWindowRecoveryService` です。
- orchestration は `final-foreground-guarantee-started` / `final-foreground-guarantee-completed` を記録し、その内部で `ExcelWindowRecoveryService.TryRecoverWorkbookWindow(..., bringToFront: true)` または active version が実行されます。

つまり foreground guarantee は decision / outcome / emit owner を orchestration に寄せ、execution primitive は `ExcelWindowRecoveryService` に残す構造です。

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
| taskpane refresh path | `TaskPaneRefreshCoordinator` | context resolve、refresh raw result |
| pane visible | `TaskPaneHostFlowService` | `taskpane-reused-shown` / `taskpane-refreshed-shown` |
| foreground guarantee decision / outcome / emit | `TaskPaneRefreshOrchestrationService` | refresh 後に走るかどうかの判断と normalized outcome / trace |
| foreground guarantee execution | `ExcelWindowRecoveryService` | app/window/foreground recovery primitive。呼び出し bridge は `TaskPaneRefreshCoordinator` |
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

## 第1実装安全単位完了後の ownership 状態

第1実装安全単位では、`CASE display completed` definition owner と `case-display-completed` emit owner を `TaskPaneRefreshOrchestrationService` へ一本化しました。これにより、already-visible path と refresh path は、同一 created-case display session の terminal 判定へ収束します。

整理済み:

- `CASE display completed` は `pane visible` の別名でも、`refresh completed` の別名でもありません。
- `TaskPaneRefreshOrchestrationService` が created-case display session を開始し、success-only の final completion を 1 回だけ成立させます。
- `WorkbookTaskPaneReadyShowAttemptWorker` は already-visible を検知して outcome を返しますが、final completion は emit しません。
- `TaskPaneRefreshCoordinator` は refresh raw facts と foreground execution result を返しますが、foreground outcome / final completion は emit しません。
- `TaskPaneHostFlowService` は pane visible state を返しますが、created-case display session 全体の terminal owner ではありません。
- `display-handoff-completed` は ready-show acceptance 側の `TaskPaneRefreshOrchestrationService` に寄せられました。

意図的に残した ownership:

- foreground guarantee owner
- visibility recovery owner
- rebuild fallback owner
- refresh source owner
- WindowActivate ownership

## 残っている ownership 論点

この節は target-state を確定するものではありません。第1実装安全単位後も current-state で分裂している ownership だけを列挙します。

### 1. ready-show orchestration owner

- current-state の分裂状態
  - request は `KernelCasePresentationService`
  - enqueue / fallback / retry 調停は `TaskPaneRefreshOrchestrationService`
  - attempt は `WorkbookTaskPaneReadyShowAttemptWorker`
- 次に設計判断が必要な点
  - ready-show protocol の owner を request から completion まで 1 本の orchestration として扱うのか。
  - fallback / retry / attempt を同一 owner に閉じるのか、それとも queue owner と attempt owner を分けたまま明示するのか。
  - ready-show が CASE display protocol のどこからどこまでを責務に含むのか。

### 2. foreground guarantee owner

- current-state の分裂状態
  - decision は `TaskPaneRefreshCoordinator`
  - execution は `ExcelWindowRecoveryService`
  - completion trace は coordinator 側にあり、mutation trace は recovery service 側にある。
- 次に設計判断が必要な点
  - foreground guarantee を 1 protocol unit として誰が ownership するのか。
  - decision と execution を分けたままにするなら、どちらを正本 owner と呼ぶのか。
  - CASE display completed との前後関係を protocol 上でどう固定するのか。

### 3. visibility recovery owner

- current-state の分裂状態
  - `WorkbookWindowVisibilityService` が workbook window visible ensure を持つ。
  - `ExcelWindowRecoveryService` が window 再解決、application visible ensure、restore、activation、foreground promotionを持つ。
  - `KernelCasePresentationService` と `WorkbookTaskPaneReadyShowAttemptWorker` がそれらを別タイミングで呼ぶ。
- 次に設計判断が必要な点
  - visibility recovery を 1 protocol としてまとめるのか、lightweight ensure と full recovery を別 protocol とするのか。
  - visible ensure と foreground promotion の境界をどこで切るのか。
  - caller 側が recovery primitive を直接組み合わせる current-state を維持するのか。

### 4. rebuild fallback owner

- current-state の分裂状態
  - rebuild fallback decision 自体は `TaskPaneSnapshotBuilderService`
  - ただし protocol 上は `TaskPaneRefreshCoordinator -> TaskPaneManager -> TaskPaneHostFlowService -> CasePaneSnapshotRenderService` の深い内部で発火する。
  - ready-show / refresh orchestration 側からは rebuild fallback が直接見えにくい。
- 次に設計判断が必要な点
  - rebuild fallback を display protocol の明示 unit として上位へ持ち上げるのか。
  - snapshot build owner と display protocol owner の境界をどこに置くのか。
  - rebuild fallback 開始時点を protocol trace 上どこで定義するのか。

### 5. refresh source owner

- current-state の分裂状態
  - refresh trigger の `reason` は caller 起点で流れる。
  - `TaskPaneRefreshCoordinator` の `refreshSource` は実質 `reason` の再掲で、source 正規化 owner がない。
  - snapshot source は `TaskPaneSnapshotBuilderService` が別に決めている。
- 次に設計判断が必要な点
  - refresh trigger source と snapshot source を同じ語で扱うのか、別概念として分けるのか。
  - refresh source の正本 owner を coordinator に置くのか、もっと上流に置くのか。
  - 観測上の `reason` と protocol 上の `source` をどう切り分けるのか。

### 6. WindowActivate 系 owner

- current-state の分裂状態
  - Excel event capture と trace は `ThisAddIn.HandleWindowActivateEvent(...)`
  - request 化と refresh dispatch は `WindowActivatePaneHandlingService`
  - `WorkbookActivate` とは別に `WorkbookLifecycleCoordinator` も refresh dispatch を持つ。
- 次に設計判断が必要な点
  - `WindowActivate` を 1 protocol unit として誰が owner になるのか。
  - event capture owner と refresh dispatch owner を分け続けるのか。
  - `WorkbookActivate` と `WindowActivate` の責務境界を protocol 上どこで分離するのか。

## foreground guarantee ownership current-state (2026-05-08 docs-only)

### current-state summary

この節は、CASE display / recovery protocol の次フェーズ候補である `foreground guarantee ownership` だけを current-state として正本化するための追記です。

- 調査開始時の `main` / `origin/main`: `3af0eb2484aa78c967b7fa5e48f252ce68907ea6`
- 第1安全単位完了記録として本文上に残る `e41feb5d607f79077e112a1945e81ac0a76d95a4` は historical completion hash として扱う。
- 参照した docs:
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
  - `docs/case-display-recovery-protocol-current-state.md`
  - `docs/case-display-recovery-protocol-target-state.md`
  - `docs/taskpane-refresh-policy.md`
  - `docs/workbook-window-activation-notes.md`
- 今回は docs-only であり、コード変更、service 分割、helper 切り出し、retry / visibility / foreground / rebuild fallback 条件変更、`WindowActivate` 挙動変更は行わない。
- docs-only のため build / test / `DeployDebugAddIn` は実行しない。

現行の foreground guarantee は、decision / outcome / terminal trace を `TaskPaneRefreshOrchestrationService` 側へ寄せ、実際の workbook window recovery / activation / foreground promotion primitive は `ExcelWindowRecoveryService` に残しています。一方で、created CASE 表示の前後には `KernelCasePresentationService`、`WorkbookTaskPaneReadyShowAttemptWorker`、`TaskPaneRefreshOrchestrationService.WorkbookPaneWindowResolver`、`CaseWorkbookOpenStrategy` にも visibility / activation / one-shot promotion が存在します。

したがって current-state では、次を分けて読む必要があります。

- `foreground guarantee`
  - refresh 後に CASE workbook / Excel window を最終 foreground へ戻す protocol unit。
- `visibility recovery`
  - workbook window を visible にする、または Excel application を visible に戻す recovery primitive。
- `window activation`
  - workbook / window / worksheet の `Activate()`、または `WorkbookActivate` / `WindowActivate` event を起点にした refresh dispatch。
- `foreground preservation`
  - hidden-for-display open 中に previous active window を戻す、または Kernel HOME を CASE より前へ戻さないための制御。
- `CASE display completed`
  - created-case display session を `TaskPaneRefreshOrchestrationService` が閉じる terminal state。

これらは似ていますが、現行コードでは同義ではありません。

### foreground guarantee 実行箇所一覧

#### 通常 path: created-case display -> ready-show -> refresh

| stage | 実行箇所 | 実行内容 | current-state の扱い |
| --- | --- | --- | --- |
| hidden-for-display open | `CaseWorkbookOpenStrategy.OpenHiddenForCaseDisplay(...)` | shared app で CASE を open し、`HideOpenedWorkbookWindow(...)` で CASE window を hidden にする。必要なら `RestorePreviousWindowForHiddenDisplay(...)` -> `RestorePreviousWindow(...)` で previous active window を `Visible = true` / `Activate()` する。 | CASE foreground guarantee ではない。表示 handoff 前の foreground preservation / flicker 抑止。 |
| initial visibility | `KernelCasePresentationService.EnsureWorkbookWindowVisibleBeforeInitialRecovery(...)` -> `WorkbookWindowVisibilityService.EnsureVisible(...)` | CASE workbook window を解決し、非表示なら `window.Visible = true`。 | visibility recovery。foreground promotion ではない。 |
| initial recovery | `KernelCasePresentationService.ShowCreatedCase(...)` -> `ExcelWindowRecoveryService.TryRecoverWorkbookWindowWithoutShowing(..., bringToFront: false)` | `ScreenUpdating` を true に戻し、window resolve / minimized restore / application visible ensure を行う。ただし `ensureWindowVisible = false`、`activateWindow = false`。 | initial recovery。foreground guarantee ではない。 |
| ready-show pre-visibility | `KernelCasePresentationService.EnsureWorkbookWindowVisibleBeforeReadyShow(...)` -> `WorkbookWindowVisibilityService.EnsureVisible(...)` | ready-show 前に CASE workbook window の visible を再確認する。 | visibility recovery。foreground promotion ではない。 |
| ready-show attempt | `WorkbookTaskPaneReadyShowAttemptWorker.TryShowWorkbookTaskPaneOnce(...)` | attempt 1 で `WorkbookWindowVisibilityService.EnsureVisible(...)` を呼ぶ。続いて `ResolveWorkbookPaneWindow(..., activateWorkbook: true)` を通し、内部で `ExcelInteropService.ActivateWorkbook(...)` が `workbook.Activate()` と first visible window `Activate()` を行う。 | window resolve / activation。final foreground guarantee ではない。 |
| refresh pre-context recovery | `TaskPaneRefreshCoordinator.TryRefreshTaskPane(...)` | Kernel HOME が visible でなければ、context 解決前に `ExcelWindowRecoveryService.TryRecoverWorkbookWindowWithoutShowing(..., bringToFront: false)` または active workbook 版を呼ぶ。 | context 解決の前提調整。foreground promotion ではない。 |
| final foreground guarantee decision | `TaskPaneRefreshOrchestrationService.CompleteForegroundGuaranteeOutcome(...)` | `refreshed && window != null && recovery service present` のときだけ foreground recovery required と判断する。 | decision / outcome owner は `TaskPaneRefreshOrchestrationService`。 |
| final foreground guarantee execution | `TaskPaneRefreshCoordinator.ExecuteFinalForegroundGuaranteeRecovery(...)` -> `ExcelWindowRecoveryService.TryRecoverWorkbookWindow(..., bringToFront: true)` または active workbook 版 | `window.Visible = true`、minimized なら `WindowState = xlNormal`、`EnsureApplicationVisible(...)`、`window.Activate()`、`PromoteWindow(...)` を実行する。`PromoteWindow(...)` は条件付き `ShowWindow(SW_RESTORE)`、topmost / no-topmost `SetWindowPos(...)`、`SetForegroundWindow(...)` を行う。 | execution owner は `ExcelWindowRecoveryService`。terminal trace は orchestration 側。 |
| post-guarantee protection | `TaskPaneRefreshCoordinator.BeginPostForegroundProtection(...)` | CASE context / workbook / window が揃う場合だけ `BeginCaseWorkbookActivateProtection(...)` を要求する。 | foreground 後の reentrant refresh 抑止。foreground guarantee の実行自体ではない。 |

#### already-visible path

| stage | 実行箇所 | 実行内容 | current-state の扱い |
| --- | --- | --- | --- |
| visible pane early-complete | `WorkbookTaskPaneReadyShowAttemptWorker.TryShowWorkbookTaskPaneOnce(...)` | `HasVisibleCasePaneForWorkbookWindow(...)` が true なら refresh を呼ばず `TaskPaneRefreshAttemptResult.VisibleAlreadySatisfied()` を返す。 | `foregroundGuaranteeRequired = false`、`foregroundGuaranteeTerminal = true` として扱われる。foreground recovery は実行されない。 |
| CASE display completed | `TaskPaneRefreshOrchestrationService.TryCompleteCreatedCaseDisplaySession(...)` | `IsPaneVisible` と `IsForegroundGuaranteeTerminal` を見て `case-display-completed` を emit する。 | CASE display completion owner は orchestration 側。already-visible path でも final foreground execution は行わない。 |

#### retry path

| retry | 実行箇所 | 実行内容 | current-state の扱い |
| --- | --- | --- | --- |
| ready retry `80ms` | `TaskPaneRefreshOrchestrationService.ScheduleTaskPaneReadyRetry(...)` | retry timer firing 後に ready-show attempt を再実行する。attempt 2 でも window resolve は `activateWorkbook: true` で走るが、`WorkbookWindowVisibilityService.EnsureVisible(...)` は attempt 1 のみ。 | retry owner は orchestration。activation は window resolve の副作用。final foreground guarantee owner ではない。 |
| pending retry `400ms` | `TaskPaneRefreshOrchestrationService.ScheduleWorkbookTaskPaneRefresh(...)` / `PendingPaneRefreshRetryService` | ready-show attempts exhausted 後、`TryRefreshTaskPane(...)` に fallback する。成功すれば通常 refresh path と同じ foreground decision / execution へ進む。 | fallback refresh owner と foreground decision は orchestration。 |
| active target fallback | `PendingPaneRefreshRetryService` -> `TaskPaneRefreshCoordinator.TryRefreshTaskPane(reason, null, null)` | 対象 workbook を見失った場合でも active CASE context があれば active refresh fallback を継続する。final guarantee は active workbook 版 `TryRecoverActiveWorkbookWindow(...)` になりうる。 | foreground target が explicit workbook ではなく active workbook へ寄る可能性がある current-state。 |

#### recovery path

| recovery | 実行箇所 | 実行内容 | current-state の扱い |
| --- | --- | --- | --- |
| lightweight visibility ensure | `WorkbookWindowVisibilityService.EnsureVisible(...)` | first visible window または `workbook.Windows[1]` を解決し、`window.Visible = true` を試みる。outcome は `AlreadyVisible` / `MadeVisible` / `WindowUnresolved` など。 | visibility owner。foreground completed とは別概念。 |
| full window recovery without showing | `ExcelWindowRecoveryService.TryRecoverWorkbookWindowWithoutShowing(...)` | `ScreenUpdating` restore、window resolve / recreate、minimized restore、application visible ensure を行う。`window.Activate()` と `PromoteWindow(...)` は行わない。 | recovery primitive。foreground guarantee ではない。 |
| full window recovery with foreground | `ExcelWindowRecoveryService.TryRecoverWorkbookWindow(..., bringToFront: true)` | full recovery に加えて `window.Activate()` と `PromoteWindow(...)` を実行する。 | final foreground guarantee の execution primitive。 |
| application foreground only | `ExcelWindowRecoveryService.TryBringApplicationToForeground(...)` | Excel application hwnd に `SetForegroundWindow(...)` を行う。 | Kernel HOME / Kernel workbook 表示系から使われる。created-case display completion の owner ではない。 |
| application show | `ExcelWindowRecoveryService.ShowApplicationWindow(...)` / `EnsureApplicationVisible(...)` | `_application.Visible = true`、`ShowWindow(SW_RESTORE)`、`ShowWindow(SW_SHOW)` を行う。 | application visibility recovery。CASE foreground guarantee と混同しない。 |

#### fallback / adjacent path

| path | 実行箇所 | 実行内容 | current-state の扱い |
| --- | --- | --- | --- |
| rebuild fallback | `TaskPaneSnapshotBuilderService` downstream of `TaskPaneRefreshCoordinator -> TaskPaneManager -> TaskPaneHostFlowService -> CasePaneSnapshotRenderService` | snapshot source が `MasterListRebuild` に落ちる場合がある。 | refresh / snapshot subprotocol。foreground guarantee owner ではない。 |
| non-`NewCaseDefault` after wait UI close | `KernelCasePresentationService.OpenCreatedCase(...)` -> `PromoteWorkbookWindowOnce(...)` | `CreateCaseSingle` など `NewCaseDefault` 以外で wait UI close 後に Excel app hwnd と workbook window hwnd を `ShowWindow(SW_RESTORE)` / topmost bounce / `SetForegroundWindow(...)` で 1 回 promote する。 | created-case display 周辺の one-shot foreground promotion。`TaskPaneRefreshCoordinator` の final foreground guarantee とは別 owner。 |
| save-before-close normalization | `KernelCaseCreationService.NormalizeInteractiveWorkbookWindowStateBeforeSave(...)` / `NormalizeBatchWorkbookWindowStateBeforeSave(...)` | hidden create session 内で save 前に owned workbook window を `Visible = true`、minimized なら `xlNormal` へ戻す。必要なら `workbook.Activate()` / `NewWindow()` で save 用 window を確保する。 | 保存状態正規化。shared/current app の CASE foreground guarantee ではない。 |
| Kernel HOME foreground | `ThisAddIn.ShowKernelHomePlaceholder(...)` / `KernelHomeForm.ForceBringToFront(...)` | `Show()`、`Activate()`、`BringToFront()`、`ShowWindow(...)`、`SetForegroundWindow(...)`、foreground retry timer を使う。 | Kernel HOME 表示 owner。CASE display foreground guarantee とは別 protocol。 |
| Kernel workbook restore / release | `KernelWorkbookDisplayService.ShowKernelWorkbookWindows(...)` / `EnsureWorkbookVisible(...)` / `ReleaseHomeDisplay(...)` | Kernel workbook window visible / normal 化、`ExcelWindowRecoveryService.TryRecoverWorkbookWindow(..., bringToFront: true|false)`、application foreground を使う。 | Kernel HOME / Kernel workbook 表示制御。CASE display foreground guarantee と混同しない。 |

### owner 分裂 / 混在ポイント

現行の owner 分裂は次のとおりです。

- foreground guarantee decision と execution が分裂している。
  - decision / terminal trace: `TaskPaneRefreshCoordinator`
  - execution primitive: `ExcelWindowRecoveryService`
- foreground terminal と recovered outcome が分裂している。
  - `GuaranteeFinalForegroundAfterRefresh(...)` は `recovered` をログに残す。
  - ただし `TaskPaneRefreshAttemptResult.Succeeded(...)` は `foregroundRecoveryStarted` をもとに terminal 化され、`recovered=false` を degraded / failed として上位へ伝えない。
- already-visible path は foreground execution を持たないが terminal として扱う。
  - `VisibleAlreadySatisfied()` は `foregroundGuaranteeRequired=false` かつ `foregroundGuaranteeTerminal=true`。
- ready-show owner と foreground owner が分裂している。
  - request: `KernelCasePresentationService`
  - enqueue / retry / fallback: `TaskPaneRefreshOrchestrationService`
  - attempt / visible already satisfied: `WorkbookTaskPaneReadyShowAttemptWorker`
  - final foreground decision: `TaskPaneRefreshOrchestrationService`
  - final foreground execution: `ExcelWindowRecoveryService`
- visibility recovery owner と foreground owner が混在しやすい。
  - `WorkbookWindowVisibilityService` は workbook window visible ensure だけを持つ。
  - `ExcelWindowRecoveryService` は application visibility、window restore、activation、foreground promotion を持つ。
  - caller は `KernelCasePresentationService`、`WorkbookTaskPaneReadyShowAttemptWorker`、`TaskPaneRefreshCoordinator` に分散している。
- WindowActivate owner が event capture と refresh dispatch に分裂している。
  - event capture: `ThisAddIn.Application_WindowActivate(...)`
  - observation / bridge: `WorkbookEventCoordinator` -> `ThisAddIn.HandleWindowActivateEvent(...)`
  - protection / suppression / refresh request: `WindowActivatePaneHandlingService`
  - actual refresh / foreground decision: `TaskPaneRefreshOrchestrationService`
- one-shot foreground promotion が final foreground guarantee と別に存在する。
  - `KernelCasePresentationService.PromoteWorkbookWindowOnce(...)` は `NewCaseDefault` 以外の created CASE 表示後に実行される。
  - これは `TaskPaneRefreshOrchestrationService` が管理する final foreground guarantee とは別 owner / 別条件で、protocol 上の `foreground guarantee completed` には統合されていない。
- Kernel HOME / Kernel workbook foreground owner が CASE display protocol と並存している。
  - `KernelHomeForm` は foreground retry timer と WinForms `Activate` / `BringToFront` / `SetForegroundWindow` を持つ。
  - `KernelWorkbookDisplayService` は Kernel workbook / Excel application の restore / foreground を持つ。
  - CASE 作成フローでは「Kernel を CASE より前に戻さない」制約もあるため、foreground preservation と foreground guarantee が混ざりやすい。
- `CASE display completed` owner とは分離済み。
  - `case-display-completed` emit owner は `TaskPaneRefreshOrchestrationService` に集約済み。
  - ただし completion 判定材料の `IsForegroundGuaranteeTerminal` は worker / coordinator が返す outcome に依存する。

### protocol 上の未定義ポイント

current-state で未定義または暗黙になっている点は次のとおりです。

- `foreground guarantee completed` の正式定義がない。
  - 実行された場合は `final-foreground-guarantee-completed` trace が観測点になる。
  - 実行されない場合は `foregroundGuaranteeTerminal=true` として上位へ返るが、`NotRequired` / `Skipped` / `NotApplicable` の明示 enum はない。
- outcome taxonomy が不足している。
  - `required` は `WasForegroundGuaranteeRequired` で表現される。
  - `completed` は `IsForegroundGuaranteeTerminal` で表現される。
  - しかし `success` / `failed` / `degraded` / `skipped` の protocol outcome は定義されていない。
  - `ExcelWindowRecoveryService` の `recovered=false` はログに残るが、CASE display completion を fail / degraded にしない。
- already-visible path の foreground obligation が暗黙。
  - visible CASE pane が既にある場合、foreground guarantee は not required と扱われる。
  - ただし「既に foreground が十分だから不要」なのか、「pane visible をもって foreground obligation を閉じる」のかは docs 上まだ明確でない。
- refresh path の foreground decision 条件が実装条件として残っている。
  - `refreshed && window != null && recoveryService != null` が current-state の required 条件。
  - この条件が UX 上の正式条件か、実装上の可能条件かは未定義。
- active workbook fallback の foreground target が未正本化。
  - workbook が null の fallback では active workbook recovery へ寄る。
  - これを created CASE display session の foreground target とみなしてよい条件は未定義。
- `WindowActivate` と foreground guarantee の関係が未定義。
  - `WindowActivate` は実 window が activate された観測点であり refresh dispatch の入口でもある。
  - ただし `WindowActivate` 発火自体を foreground guarantee completed とみなすかどうかは未定義。
- one-shot promotion と final foreground guarantee の関係が未定義。
  - `PromoteWorkbookWindowOnce(...)` と `GuaranteeFinalForegroundAfterRefresh(...)` は別々に存在する。
  - どちらが CASE display protocol の foreground owner なのか、または別 unit なのかは未正本化。
- rebuild fallback と foreground guarantee の関係は明示が必要。
  - current-state では rebuild fallback は snapshot / refresh 内部の話で、foreground guarantee owner ではない。
  - ただし refresh が成功した後に foreground decision が走るため、観測上は同一 refresh attempt 内に見える。

### 守るべき既存制約

foreground guarantee target-state 化では、次の現行制約を壊さないことを前提にする。

- 白Excel対策
  - `PostCloseFollowUpScheduler` の no visible workbook quit 設計目標と、Excel application visibility recovery を混同しない。
  - foreground owner 整理のために白 Excel を覆うだけの追加ガードを足さない。
- COM解放
  - hidden create session / retained hidden app-cache / temporary workbook close の既存 cleanup 境界を変えない。
  - foreground guarantee の owner 整理を理由に COM 参照の lifetime を広げない。
- Excel状態制御
  - `ScreenUpdating` / `DisplayAlerts` / `EnableEvents` は既存 scope で restore する。
  - `ExcelWindowRecoveryService.EnsureScreenUpdatingEnabled(...)` は recovery primitive として扱い、shared state の恒常変更にしない。
- fail closed
  - context / workbook / window が解決できない場合に推測で補完しない。
  - foreground target が不明なまま active workbook promotion へ広げる変更をしない。
- timing hack 禁止
  - `Application.DoEvents()`、sleep、単なる delay 追加で foreground guarantee を定義しない。
  - retry 値や timer 条件は今回変更しない。
- ガード追加で覆わない
  - foreground / visibility / rebuild fallback 条件を新しい guard で覆って挙動を隠さない。
  - `WorkbookOpen` 直後を window 安定境界へ戻さない。

### 次に target-state 化すべき論点

次フェーズで target-state 化する場合、少なくとも次を先に決める必要があります。

1. `foreground guarantee completed` を outcome enum として定義するか。
   - 例: `RequiredSucceeded`、`RequiredFailedDegraded`、`NotRequired`、`SkippedByAlreadyVisible`、`SkippedNoWindow` など。
   - ただしこれは target-state 論点であり、current-state では未定義。
2. foreground guarantee completion owner は `TaskPaneRefreshOrchestrationService` とし、`TaskPaneRefreshCoordinator` は execution bridge として残す。
3. `ExcelWindowRecoveryService` の `recovered=false` を display completion に影響させるか、観測ログに留めるか。
4. already-visible path の foreground obligation をどう閉じるか。
5. `PromoteWorkbookWindowOnce(...)` を CASE display protocol の foreground unit に含めるか、non-`NewCaseDefault` の historical one-shot promotion として別扱いするか。
6. WindowActivate event と foreground guarantee の関係をどう定義するか。
7. active workbook fallback の foreground target を、created CASE session と紐づけてよい条件を明文化するか。
8. rebuild fallback は refresh / snapshot subprotocol に留め、foreground / CASE display completion 条件へ昇格させないことを target-state でも固定するか。

### 今回行わないこと

- コード変更なし。
- service 分割なし。
- helper 切り出しなし。
- retry 条件変更なし。
- visibility recovery 条件変更なし。
- foreground recovery 条件変更なし。
- rebuild fallback 条件変更なし。
- `WindowActivate` 挙動変更なし。
- build / test / `DeployDebugAddIn` 実行なし。

## visibility recovery ownership current-state (2026-05-08 docs-only)

### current-state summary

この節は、現行 `main` の visibility recovery owner を current-state として正本化するための docs-only 追記です。

- 調査開始時の `main` / `origin/main`: `82d125567085220e6998c124882df9fba31e095c`
- 参照した docs:
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
  - `docs/taskpane-refresh-policy.md`
  - `docs/case-display-recovery-protocol-current-state.md`
  - `docs/case-display-recovery-protocol-target-state.md`
- 今回は visibility recovery の current-state 把握と正本化だけを行う。
- コード変更、service 分割、helper 切り出し、ready retry / pending retry / visibility recovery / foreground guarantee / rebuild fallback / `WindowActivate` 条件変更は行わない。
- docs-only のため build / test / `DeployDebugAddIn` は実行しない。

current-state の visibility recovery は、単一 owner ではなく次の概念に分かれている。

- workbook window visible ensure
  - `WorkbookWindowVisibilityService.EnsureVisible(...)` が lightweight primitive owner。
  - `ExcelInteropService.GetFirstVisibleWindow(workbook)`、fallback の `workbook.Windows[1]`、`Window.Visible` 読み取り / `window.Visible = true` によって workbook window の visible を扱う。
- full Excel window recovery
  - `ExcelWindowRecoveryService` が full primitive owner。
  - `ScreenUpdating` restore、window resolve / recreate、minimized restore、application visible ensure、必要時の activation / foreground promotion を扱う。
- pane visible state
  - `TaskPaneDisplayCoordinator` / `TaskPaneHostFlowService` が host/pane 側の visible state owner。
  - `WorkbookWindowVisibilityService` の `AlreadyVisible` / `MadeVisible` は workbook window visibility であり、CASE pane visible とは同義ではない。

したがって、visibility recovery の current-state は「誰が recovery primitive を持つか」と「誰がその primitive を呼ぶか」と「pane visible を誰が判定するか」が分裂した状態である。

### pane visible 判定の current-state

`pane visible` は、現行コードでは protocol-level の独立 outcome ではなく、host metadata と VSTO CustomTaskPane の visible 状態の join として判定される。

- already-visible path の判定入口は `WorkbookTaskPaneReadyShowAttemptWorker.TryShowWorkbookTaskPaneOnce(...)`。
- worker は `ResolveWorkbookPaneWindow(..., activateWorkbook: true)` で window を解決した後、`HasVisibleCasePaneForWorkbookWindow(...)` を呼ぶ。
- 実体は `ThisAddIn -> TaskPaneManager -> TaskPaneDisplayCoordinator.HasVisibleCasePaneForWorkbookWindow(...)`。
- 判定条件は次の組み合わせである。
  - window key が解決できる。
  - `_hostsByWindowKey` に対象 window key の host がある。
  - host の `WorkbookFullName` が対象 workbook と一致する。
  - host role が `WorkbookRole.Case` である。
  - `TaskPaneHost.IsVisible` が true である。
- `TaskPaneHost.IsVisible` は内部の `CustomTaskPane.Visible` を読む。
- refresh path で pane が show された事実は `TaskPaneHostFlowService` が `taskpane-reused-shown` / `taskpane-refreshed-shown` と `TaskPaneHostFlowResult.ReusedShown()` / `RefreshedShown()` で返す。

このため current-state では、`pane visible` は `WorkbookWindowVisibilityService.EnsureVisible(...)` の outcome ではない。workbook window が visible でも pane host が未生成、別 workbook、別 window key、非 CASE、または `CustomTaskPane.Visible=false` なら visible CASE pane とは判定されない。

### visibility recovery 実行箇所 / 判定箇所一覧

| stage | 実行 / 判定箇所 | 実行内容 | current-state の扱い |
| --- | --- | --- | --- |
| hidden-for-display open | `CaseWorkbookOpenStrategy.OpenHiddenForCaseDisplay(...)` | created CASE を shared app で open し、opened workbook window を hidden にする。必要なら previous active window を `Visible = true` / `Activate()` で戻す。 | visibility recovery というより display handoff 前の foreground preservation / flicker 抑止。CASE pane visible owner ではない。 |
| initial workbook-window visibility | `KernelCasePresentationService.EnsureWorkbookWindowVisibleBeforeInitialRecovery(...)` -> `WorkbookWindowVisibilityService.EnsureVisible(...)` | initial recovery 前に CASE workbook window を解決し、非表示なら `window.Visible = true` を試みる。 | lightweight workbook window visibility recovery。 |
| initial full recovery without showing | `KernelCasePresentationService.ShowCreatedCase(...)` -> `ExcelWindowRecoveryService.TryRecoverWorkbookWindowWithoutShowing(..., bringToFront: false)` | `ScreenUpdating` restore、window resolve / recreate、minimized restore、application visible ensure を行う。`ensureWindowVisible=false`、`activateWindow=false`。 | full recovery primitive だが foreground guarantee ではない。 |
| ready-show pre-visibility | `KernelCasePresentationService.EnsureWorkbookWindowVisibleBeforeReadyShow(...)` -> `WorkbookWindowVisibilityService.EnsureVisible(...)` | ready-show request 前に CASE workbook window visibility を再確認する。 | ready-show 前の lightweight recovery。 |
| ready-show attempt pre-visibility | `WorkbookTaskPaneReadyShowAttemptWorker.EnsureWorkbookWindowVisibleForTaskPaneDisplay(...)` -> `WorkbookWindowVisibilityService.EnsureVisible(...)` | attempt 1 のみ workbook window visible ensure を行う。attempt 2 ではこの ensure は走らない。 | ready-show attempt 内の lightweight recovery。ready retry owner ではない。 |
| ready-show window resolve / activation | `WorkbookTaskPaneReadyShowAttemptWorker.TryShowWorkbookTaskPaneOnce(...)` -> `TaskPaneRefreshOrchestrationService.ResolveWorkbookPaneWindow(..., activateWorkbook: true)` | `ExcelInteropService.ActivateWorkbook(...)` が `workbook.Activate()` と first visible window `Activate()` を行い、visible window または active workbook/window fallback を返す。 | window resolve / activation。visibility recovery completed ではない。 |
| already-visible pane 判定 | `TaskPaneDisplayCoordinator.HasVisibleCasePaneForWorkbookWindow(...)` | window key、host、workbook fullname、CASE role、`CustomTaskPane.Visible` を照合する。 | pane visible 判定。refresh trigger ではなく early-complete 判定材料。 |
| refresh pre-context recovery | `TaskPaneRefreshCoordinator.TryRefreshTaskPane(...)` | Kernel HOME が visible でなく `ExcelWindowRecoveryService` が存在する場合、context 解決前に `TryRecoverWorkbookWindowWithoutShowing(...)` または active workbook 版を呼ぶ。 | context 解決の前提調整。foreground promotion ではない。 |
| refresh target window resolve | `TaskPaneRefreshCoordinator.EnsurePaneWindowForWorkbook(...)` -> `ResolveWorkbookPaneWindow(..., activateWorkbook: false)` | workbook 指定時に pane 対象 window を確定する。 | refresh path の window target 補完。 |
| pane show / visible state | `TaskPaneHostFlowService.TryReuseCaseHostForRefresh(...)` / `RenderAndShowHostForRefresh(...)` | `TaskPaneDisplayCoordinator.TryShowHost(...)` で host を show し、`taskpane-reused-shown` / `taskpane-refreshed-shown` を記録する。 | actual pane visible transition の owner。 |
| ready-show fallback handoff | `TaskPaneRefreshOrchestrationService.ScheduleWorkbookTaskPaneRefresh(...)` / `PendingPaneRefreshRetryService` | ready-show attempts exhausted 後、workbook target または active target を追って refresh retry へ移る。 | retry / fallback owner。visibility primitive owner ではない。 |
| final foreground guarantee | `TaskPaneRefreshOrchestrationService.CompleteForegroundGuaranteeOutcome(...)` -> `TaskPaneRefreshCoordinator.ExecuteFinalForegroundGuaranteeRecovery(...)` -> `ExcelWindowRecoveryService.TryRecoverWorkbookWindow(..., bringToFront: true)` | refresh success 後に必要条件が揃った場合だけ full recovery + activation + foreground promotion を実行する。 | foreground guarantee。visibility recovery とは境界を分ける。 |
| application foreground / show adjacent path | `ExcelWindowRecoveryService.TryBringApplicationToForeground(...)` / `ShowApplicationWindow(...)` / `EnsureApplicationVisible(...)` | Excel application hwnd の foreground request、application visible 化、`ShowWindow` を行う。 | Kernel HOME / Kernel workbook など adjacent owner からも使われる primitive。created CASE pane visible owner ではない。 |

### owner 分裂 / 混在ポイント

現行の visibility recovery ownership は次のように分裂している。

- primitive owner が 2 層に分かれている。
  - lightweight workbook window visible ensure: `WorkbookWindowVisibilityService`
  - full application / workbook window recovery: `ExcelWindowRecoveryService`
- caller orchestration owner が複数に分かれている。
  - pre-handoff / post-release request 前: `KernelCasePresentationService`
  - ready-show attempt: `WorkbookTaskPaneReadyShowAttemptWorker`
  - refresh pre-context recovery: `TaskPaneRefreshCoordinator`
  - retry / fallback scheduling と foreground outcome: `TaskPaneRefreshOrchestrationService`
- pane visible 判定 owner は recovery primitive owner と別である。
  - already-visible 判定は `TaskPaneDisplayCoordinator.HasVisibleCasePaneForWorkbookWindow(...)`
  - actual show result は `TaskPaneHostFlowService`
  - `WorkbookWindowVisibilityService.EnsureVisible(...)` は CASE pane visible を返さない。
- ready-show retry owner と visibility primitive owner が混ざりやすい。
  - ready retry `80ms` は `TaskPaneRefreshOrchestrationService.ScheduleTaskPaneReadyRetry(...)` が担う。
  - attempt 1 の `WorkbookWindowVisibilityService.EnsureVisible(...)` は worker が呼ぶが、retry 継続可否そのものは visibility outcome ではなく attempt success / attempts count に従う。
- pending retry owner と visibility target が混ざりやすい。
  - pending retry `400ms` は `PendingPaneRefreshRetryService` が workbook target / active target を追う。
  - workbook target を見失うと active CASE context fallback に寄るため、後続 recovery target は explicit workbook ではなく active workbook になりうる。
- foreground guarantee owner とは分離済みだが、primitive は重なる。
  - foreground guarantee の decision / normalized outcome / trace owner は `TaskPaneRefreshOrchestrationService`。
  - execution primitive は `ExcelWindowRecoveryService`。
  - `ExcelWindowRecoveryService` は visibility recovery と foreground promotion の両方を持つため、呼び出し条件と outcome owner を混同しやすい。
- rebuild fallback owner とは別である。
  - rebuild fallback は `TaskPaneSnapshotBuilderService` の snapshot acquisition subprotocol。
  - visibility recovery が失敗したから rebuild fallback へ入る、という直接接続ではない。
- `CASE display completed` owner とは分離済みである。
  - `TaskPaneRefreshOrchestrationService.TryCompleteCreatedCaseDisplaySession(...)` は `IsPaneVisible` と `ForegroundGuaranteeOutcome.IsTerminal / IsDisplayCompletable` を消費する。
  - workbook window visibility ensure の outcome をそのまま `case-display-completed` に使わない。
- `WindowActivate` owner とは別である。
  - event capture は `ThisAddIn` / `WorkbookEventCoordinator`。
  - protection / suppression / refresh dispatch は `WindowActivatePaneHandlingService`。
  - actual refresh / foreground outcome は `TaskPaneRefreshOrchestrationService`。
  - `WindowActivate` 発火だけを visibility recovery completed とは扱わない。

### protocol 上の未定義ポイント

current-state で未定義または暗黙になっている visibility recovery 論点は次のとおりです。

- `visibility recovery completed / skipped / degraded / failed` 相当の protocol-level outcome は定義されていない。
  - `WorkbookWindowVisibilityEnsureOutcome` は lightweight primitive の local outcome であり、display protocol 全体の outcome ではない。
  - `ExcelWindowRecoveryService` は bool を返すが、`without showing` と `with foreground` の意味差や recovered field の扱いは visibility outcome として正規化されていない。
- 何をもって `pane visible` とするかは実装条件として存在するが、protocol 定義として独立していない。
  - 現行判定は window key + host registry + workbook fullname + `WorkbookRole.Case` + `CustomTaskPane.Visible` の join。
  - この join が `CASE display completed` の hard requirement としてどう命名されるかは未定義。
- workbook window visible と pane visible の関係が暗黙である。
  - workbook window visible は pane visible の前提になりうるが、十分条件ではない。
  - `AlreadyVisible` / `MadeVisible` は CASE pane visible を意味しない。
- visibility recovery 失敗と retry / fallback の接続が明示されていない。
  - ready-show attempt では window resolve / refresh success が最終的な attempt success を決める。
  - `WorkbookWindowVisibilityService.EnsureVisible(...)` の `WindowUnresolved` / `Failed` が直接 pending retry や rebuild fallback を選ぶわけではない。
- ready retry attempt 2 の visibility ensure 方針は current-state では実装事実に留まる。
  - attempt 2 でも `ResolveWorkbookPaneWindow(..., activateWorkbook: true)` は走る。
  - ただし `WorkbookWindowVisibilityService.EnsureVisible(...)` は attempt 1 のみであり、この差の UX 上の正式意図は未定義。
- pending retry active target fallback の visibility target が未正本化である。
  - active CASE context fallback に入った場合、created-case display session の target と active workbook recovery target を同一視してよい条件は未定義。
- `WindowActivate` と visibility recovery completed の関係が未定義である。
  - `WindowActivate` は observed activation / refresh dispatch の入口であり、pane visible や visibility recovery completed の代替ではない。
- `VisibleAfterSet=false/null` や full recovery の partial failure を degraded / failed として上位へ伝える protocol がない。
  - 現行ではログ観測に留まり、CASE display completion の success-only 判定は pane visible と foreground outcome 側で閉じる。

### rebuild fallback との接続点

rebuild fallback は visibility recovery の前段条件ではなく、refresh path に入った後の snapshot acquisition subprotocol である。

- already-visible path が成立した場合、refresh 自体を呼ばないため rebuild fallback には入らない。
- already-visible path が成立しない場合、worker は `TaskPaneRefreshCoordinator.TryRefreshTaskPane(...)` へ handoff する。
- refresh path で context が受理され、`TaskPaneManager.RefreshPaneWithOutcome(...)` に進んだ後、`TaskPaneHostFlowService -> TaskPaneManager.RenderCaseHost(...) -> CasePaneSnapshotRenderService.Render(...) -> TaskPaneSnapshotBuilderService.BuildSnapshotText(...)` の順で snapshot source decision に入る。
- `TaskPaneSnapshotBuilderService` が `CaseCache -> BaseCache -> BaseCacheFallback -> MasterListRebuild` を選ぶ。
- `MasterListRebuild` は pane 内容を構築するための fallback であり、workbook window visibility や Excel foreground を回復する owner ではない。
- rebuild fallback が発生しても、`CASE display completed` の直接条件にはしない。refresh が pane visible を成立させ、その後 foreground obligation が terminal になった場合だけ completion 材料になる。
- 逆に window resolve / context 解決 / refresh precondition が fail-closed で止まる場合、rebuild fallback まで到達しない。

### foreground guarantee との境界

visibility recovery と foreground guarantee は、`ExcelWindowRecoveryService` という primitive を共有しうるが、protocol unit と owner は別である。

- `WorkbookWindowVisibilityService.EnsureVisible(...)` は workbook window visible ensure だけを扱う。
- `ExcelWindowRecoveryService.TryRecoverWorkbookWindowWithoutShowing(..., bringToFront: false)` は `ensureWindowVisible=false` / `activateWindow=false` で使われる initial / pre-context recovery であり、foreground guarantee ではない。
- `ExcelWindowRecoveryService.TryRecoverWorkbookWindow(..., bringToFront: true)` は final foreground guarantee execution primitive として使われる。
- foreground guarantee の decision / normalized outcome / `foreground-recovery-decision` / `final-foreground-guarantee-*` trace owner は `TaskPaneRefreshOrchestrationService`。
- already-visible path では `ForegroundGuaranteeOutcome.SkippedAlreadyVisible(...)` 相当で foreground execution は走らない。
- `Window.Visible = true` が実行されたこと、または `WindowActivate` が発火したことだけでは foreground guarantee completed とは扱わない。

### WindowActivate / WorkbookOpen との境界

- `WorkbookOpen` は window 安定境界ではない。
  - `TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(...)` は `reason == WorkbookOpen` かつ workbook あり / window なしを skip する shared policy。
  - visibility recovery owner 整理を理由に、この境界を戻さない。
- `WindowActivate` は event capture と refresh dispatch に分かれる。
  - capture: `ThisAddIn.Application_WindowActivate(...)`
  - observation bridge: `WorkbookEventCoordinator.OnWindowActivate(...)` -> `ThisAddIn.HandleWindowActivateEvent(...)`
  - protection / suppression / dispatch: `WindowActivatePaneHandlingService`
  - refresh / foreground outcome: `TaskPaneRefreshOrchestrationService`
- `WindowActivate` は visible window が activate された観測点であり、pane visible 判定、visibility recovery outcome、foreground guarantee completed の代替ではない。
- protection / suppression により `WindowActivate` dispatch が return する場合があるため、event 発火だけで refresh path 到達を保証しない。

### 守るべき既存制約

visibility recovery target-state 化では、次の current-state 制約を壊さないことを前提にする。

- 白Excel対策
  - `PostCloseFollowUpScheduler` の no visible workbook quit と visibility recovery を混同しない。
  - 白 Excel を覆うだけの追加ガードで recovery 条件を隠さない。
- TaskPane が出ない regression の防止
  - ready-show、already-visible early-complete、pending retry、host reuse / render / show の現行条件を変更しない。
- COM解放
  - hidden create session、retained hidden app-cache、一時 workbook close の cleanup 境界を変えない。
  - visibility outcome 整理のために workbook / window / application COM reference lifetime を広げない。
- Excel状態制御
  - `ScreenUpdating` / `DisplayAlerts` / `EnableEvents` の既存 restore scope を変えない。
  - `ExcelWindowRecoveryService.EnsureScreenUpdatingEnabled(...)` を恒常状態変更として扱わない。
- fail closed
  - workbook / window / context が不明な場合に推測で補完しない。
  - active workbook fallback を target 不明時の広域 promotion として拡大しない。
- timing hack 禁止
  - `Application.DoEvents()`、sleep、単なる delay 追加で visibility completed を作らない。
  - ready retry `80ms`、pending retry `400ms`、attempt count は今回変更しない。
- ガード追加で覆わない
  - visibility / foreground / rebuild fallback 条件を新しい guard で隠さない。
  - `WorkbookOpen` を window 安定境界へ戻さない。

### 次に target-state 化すべき論点

次フェーズで visibility recovery を target-state 化する場合、少なくとも次を先に決める必要がある。

1. `visibility recovery` を protocol unit として定義するか。
   - workbook window visible ensure、full window recovery、pane visible を 1 unit にまとめるのか、別 unit として残すのか。
2. `visibility recovery completed / skipped / degraded / failed` 相当の normalized outcome を定義するか。
   - `WorkbookWindowVisibilityEnsureOutcome` をそのまま上位 outcome にしない場合、どこで変換するか。
3. `pane visible` の canonical definition を docs 上で固定するか。
   - window key、host registry、workbook fullname、CASE role、`CustomTaskPane.Visible` の join を正式条件として扱うか。
4. visibility recovery caller orchestration owner をどこへ寄せるか。
   - pre-handoff は `KernelCasePresentationService` に残し、post-handoff は `TaskPaneRefreshOrchestrationService` に寄せる方針を採るか。
5. ready-show attempt 1 のみ lightweight ensure を行う現行条件を、target-state でどう説明するか。
6. visibility recovery failure を ready retry / pending retry / fail-closed にどう接続するか。
   - rebuild fallback へ直接接続しないことを target-state でも固定するか。
7. pending retry の active CASE context fallback 時に、visibility / foreground target をどう正本化するか。
8. `WindowActivate` を visibility recovery の観測点としてだけ扱うのか、refresh request source としてだけ扱うのか。
9. `CASE display completed` が消費するのは `pane visible` と foreground terminal であり、workbook window visible ensure outcome ではないことを target-state でも固定するか。

### 今回行わないこと

- コード変更なし。
- service 分割なし。
- helper 切り出しなし。
- ready retry 条件変更なし。
- pending retry 条件変更なし。
- visibility recovery 条件変更なし。
- foreground guarantee 条件変更なし。
- rebuild fallback 条件変更なし。
- `WindowActivate` 挙動変更なし。
- build / test / `DeployDebugAddIn` 実行なし。

## 今回の current-state に含めないこと

- 表示不安定の原因断定
- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` の timing 変更
- ready-show retry 値の調整
- foreground recovery 条件の変更
- visibility 制御の変更
- hidden session 修正
- CASE 作成本体の修正

## 一言まとめ

現行の CASE display / recovery protocol は、単一 owner の直線的な flow ではありません。第1実装安全単位で display completion は `TaskPaneRefreshOrchestrationService` に集約されましたが、ready-show、refresh、foreground guarantee、visibility recovery、rebuild fallback、WindowActivate は複数 service に分散したままです。

この文書では、completion definition を整理済みの到達点として固定し、残りの ownership を次の安全単位候補として分けて扱います。
