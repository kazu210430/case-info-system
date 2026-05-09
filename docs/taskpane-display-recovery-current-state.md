# TaskPane 表示回復 Current State

## 位置づけ

この文書は、TaskPane 表示回復領域の Phase 0 正本です。

目的は巨大クラスを小さくすることではありません。現行フロー、route、retry、ready-show、foreground outcome、window resolve、trace、fail-closed 条件、変更禁止順序、現在の orchestration 集中箇所、本来あるべき責務境界を 1 つの読み方に固定し、次フェーズで安全単位を選べる状態にすることです。

参照した正本:

- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/taskpane-refresh-policy.md`
- `docs/taskpane-display-recovery-freeze-line.md`
- `docs/taskpane-refresh-orchestration-target-boundary-map.md`
- `docs/current-flow-source-of-truth.md`
- `docs/case-display-recovery-protocol-current-state.md`
- `docs/case-display-recovery-protocol-target-state.md`
- `docs/visibility-foreground-boundary-current-state.md`
- `docs/readyshow-recovery-observation-points-2026-05-08.md`

この文書は docs-only です。コード移動、service 追加、retry 値変更、trace 名変更、route 契約変更、fail-closed 条件変更、window activation policy 変更、COM restore 順序変更、UI policy 変更は行いません。

## 全工程での位置

| Phase | 内容 | この文書との関係 |
| --- | --- | --- |
| Phase 0 | 正本 docs を固定する | 今回の対象。TaskPane 表示回復を current-state として固定する。 |
| Phase 1 | 現行巨大クラスの responsibility inventory を作る | `docs/taskpane-refresh-orchestration-responsibility-inventory.md` で `TaskPaneRefreshOrchestrationService` の responsibility inventory を固定する。 |
| Phase 2 | 理想責務境界との対応表を作る | `docs/taskpane-refresh-orchestration-target-boundary-map.md` で、この文書の「本来あるべき責務境界」を R01-R16 の対応表へ展開する。 |
| Phase 3 | 変更禁止領域を固定する | `docs/taskpane-display-recovery-freeze-line.md` を正本として、この文書の「変更禁止順序」と「fail-closed 条件」を freeze line へ展開する。 |
| Phase 4 | safe-first ownership separation を閉じる | route / retry / trace / order を変えず、semantics owner ではない層の owner separation を完了扱いにする。これ以上の runtime extraction は Phase 5 の protocol-preserving redesign 前に行わない。 |
| Phase 5 | protocol-preserving convergence redesign として orchestration shrink を扱う | completion convergence、callback meaning、display session ownership、foreground linkage を同時に扱う。単なる class shrink として始めない。 |

## Phase 4 closure current-state

Phase 4 safe-first ownership separation はここで終了扱いです。

Phase 4 の目的は「巨大クラスを小さくする」ことではなく、次の 4 点でした。

- ownership separation。
- lifecycle visibility。
- danger boundary localization。
- freeze line stabilization。

safe-first で切れる runtime owner は概ね分離済みです。具体的には、R02 refresh precondition / fail-closed policy、R16 timer lifecycle、R06 ready-show retry scheduler、R10/R11/R12 normalized outcome mapping、R15 WindowActivate downstream observation、R08 pending retry owner file boundary separation が Phase 4 の runtime separation として完了済みです。

また、Phase 4 終了前の docs freeze として、ready-show retry contract truth、R07 pending fallback semantics、R09 window resolver / `activateWorkbook` route matrix、active fallback truth table、foreground outcome contract を固定済みです。これらは Phase 5 で protocol core に触る前の前提として扱います。

残件は safe-first extraction ではなく、display recovery protocol の核心領域です。

- completion convergence。
- display session ownership。
- foreground linkage。
- protocol entry meaning。
- retry convergence。
- callback-to-completion bridge。

したがって、これ以上 Phase 4 として runtime extraction を続けません。追加 extraction は completion semantics、callback meaning、retry sequencing、display session boundary、foreground outcome semantics、trace contract に近づき、safe-first ではなく protocol rewrite になりやすいためです。

Phase 5 の入口は、単なる「orchestration shrink」ではなく、protocol-preserving convergence redesign として扱います。最初に設計単位として扱う候補は、R05 callback/completion convergence、R10/R13/R14 foreground + completion convergence、display protocol convergence map です。

現時点で runtime extraction STOP とする領域:

- R04/R14 display session。
- R05 callback/completion convergence。
- R07 fallback handoff。
- R09 window resolver。
- R13 foreground linkage。

immutable freeze line:

- attempt 1 -> 80ms attempt 2 -> pending fallback。
- pending retry 400ms / 3 attempts。
- pending != completion。
- WindowActivate dispatch != completion。
- case-display-completed one-time emit。
- display session boundary。
- trace contract。
- callback meaning。
- retry sequencing。
- foreground outcome semantics。

## Phase 5 display protocol convergence contract

Phase 5 の開始点として、display protocol convergence contract を docs-only で固定します。

Phase 5 は「巨大 orchestration を分割する」ことを目的にしません。目的は、現在の runtime protocol を保ったまま、completion ownership、callback meaning、foreground linkage、retry convergence、display session boundary を同じ読み方へ収束させることです。

今回固定する convergence topology:

```text
raw facts
↓
R10/R11/R12 normalization
↓
R13 foreground interpretation
↓
R14 completion gate
↓
one-time emit
```

この topology の読み方:

- raw facts は ready-show attempt、already-visible path、refresh path、pending retry、active CASE fallback、WindowActivate downstream observation、foreground execution raw result から来る観測事実です。
- R10/R11/R12 は visibility / refresh source / rebuild fallback を completion 判定可能な normalized outcome に変換します。ただし completion owner ではありません。
- R13 は foreground outcome を completion input として解釈します。ただし foreground outcome chain 自体は emit owner ではありません。
- R14 だけが completion gate であり、created CASE display session の one-time emit owner です。
- one-time emit は `case-display-completed` だけです。worker、retry、WindowActivate、foreground、fallback、host-flow へ emit ownership を戻しません。

### Callback meaning freeze

ready-show callback は、`shown raw facts` を orchestration convergence chain に戻す callback です。completion callback ではありません。

読み替え禁止:

- ready-show callback = display completed。
- ready-show callback = recovery completed。
- ready-show callback = foreground completed。
- ready-show callback = final success。

`WorkbookTaskPaneReadyShowAttemptWorker` は attempt を実行し、`WorkbookTaskPaneReadyShowAttemptOutcome` を返します。`TaskPaneDisplayRetryCoordinator` は attempt sequencing を扱い、shown 時に callback を呼びます。ただし callback 後も、visibility outcome、refresh source outcome、rebuild fallback outcome、foreground outcome、created-case display session の completion gate を通るまで `case-display-completed` ではありません。

### Non-completion events contract

次の trace / event / result は completion に見えやすいですが、completion ではありません。いずれも `case-display-completed` の直接代替にしません。

| event / result | completion ではない理由 |
| --- | --- |
| `taskpane-already-visible` | visible host の raw fact です。already-visible path の success 相当 fact であって、R10/R11/R12/R13/R14 を通る前の completion ではありません。 |
| `taskpane-refresh-completed` | refresh execution の完了観測です。pane visible / visibility terminal / foreground terminal / display session one-time gate を満たすまでは completion ではありません。 |
| `foreground-recovery-decision` | foreground outcome の decision trace です。completion input にはなり得ますが emit owner ではありません。 |
| `final-foreground-guarantee-completed` | foreground execution bridge の完了観測です。`RequiredSucceeded` / `RequiredDegraded` 等へ正規化され、R14 gate を通るまでは completion ではありません。 |
| `display-refresh-trigger-dispatched` | WindowActivate 由来の downstream refresh entry へ渡した観測です。display success ではありません。 |
| `window-activate-display-refresh-trigger-outcome` | WindowActivate downstream refresh の観測です。completion trace ではありません。 |
| `defer-retry-end refreshed=true` | pending retry の refresh attempt result です。retry owner は convergence owner ではありません。 |
| `defer-active-context-fallback-end refreshed=true` | target-lost resiliency fallback の refresh attempt result です。active fallback success は completion ではありません。 |
| `resolve-window-success` 相当の window resolve success | window availability fact です。foreground success、retry success、display completed のいずれでもありません。 |
| `ready-show-attempt-result refreshed=true` | attempt raw result です。callback 後の normalized outcome chain と completion gate を満たすまで completion ではありません。 |
| `RequiredSucceeded` foreground outcome | display-completable terminal input です。foreground success を direct completion と読みません。 |
| `RequiredDegraded` foreground outcome | display-completable terminal input です。success / failure へ丸めず、direct completion と読みません。 |

completion trace は `case-display-completed` だけです。

### Foreground と completion の距離

foreground outcome は completion input ですが、foreground outcome chain 自体は completion owner ではありません。

固定する意味:

- foreground decision / execution / outcome は R13 の interpretation です。
- R13 は R14 の input を作れますが、R14 を代替しません。
- `RequiredDegraded` は display-completable terminal として扱います。
- `RequiredDegraded` を success / failure へ丸めません。
- `RequiredDegraded` を completion へ直接読み替えません。

### Pending / active fallback と completion の距離

pending retry success、active fallback success は completion ではありません。

固定する意味:

- `PendingPaneRefreshRetryService` は pending retry `400ms / 3 attempts`、tracked workbook retry、active CASE context fallback の owner です。
- pending retry owner は convergence owner ではありません。
- active CASE fallback は target-lost resiliency fallback です。completion fallback ではありません。
- retry success / fallback success は raw fact として existing refresh / outcome / completion chain に戻るだけです。
- `refreshed=true` を recovered、foreground success、display completed、`case-display-completed` のいずれにも読み替えません。

### One-time emit owner

`case-display-completed` の one-time emit owner は orchestration convergence owner だけです。

固定する owner:

- `TaskPaneRefreshOrchestrationService` が current owner です。
- R14 completion gate が created CASE display session の completion hard gate です。
- worker、retry、WindowActivate、foreground、fallback、host-flow、TaskPaneManager へ ownership を戻しません。

completion hard gate は、少なくとも次をすべて必要とします。

- created CASE display reason であること。
- refresh attempt result が存在すること。
- refresh success であること。
- pane visible であること。
- visibility outcome が terminal かつ display-completable であること。
- foreground guarantee が terminal であること。
- foreground outcome が display-completable であること。
- created CASE display session が解決できること。
- session が未完了であること。

この hard gate を満たした場合だけ、session ごとに 1 回だけ `case-display-completed` を emit します。

### Phase 5 boundary

Phase 5 は runtime extraction を急ぐフェーズではありません。

優先すること:

- convergence topology clarification。
- ownership preservation。
- protocol-preserving redesign。
- completion ownership clarification。
- callback meaning clarification。
- display session convergence mapping。

禁止する読み方:

- R10/R11/R12 normalized outcome を completion owner とみなす。
- R13 foreground outcome を emit owner とみなす。
- pending / active fallback success を completion とみなす。
- WindowActivate dispatch / downstream observation を completion とみなす。
- ready-show callback を completion callback とみなす。
- one-time emit owner を lower-level worker / retry / foreground / WindowActivate へ分散する。

### Phase 5 second runtime state

Phase 5 第二実装後の runtime refactor 完了点は次の 2 点です。

- completion hard gate yes/no decision の private helper 化。
- `case-display-completed` details payload assembly の private helper 化。

どちらも `TaskPaneRefreshOrchestrationService` 内の R14 orchestration に留まる局所 helper 化です。owner 移動ではありません。

R14 owner preserved:

- completion owner は未移動です。
- `case-display-completed` emit owner は未移動です。
- one-time emit guard は未移動です。
- display session lookup は未移動です。
- `IsCompleted` guard は未移動です。
- lock は未移動です。
- dictionary remove は未移動です。
- `NewCaseVisibilityObservation.Complete(...)` は未移動です。
- trace emit position は未移動です。

helper の意味:

- hard gate decision helper は、visibility / foreground display-completable facts に基づく yes/no decision だけを担います。
- payload helper は、`case-display-completed` details payload assembly だけを担います。
- どちらの helper も completion emit しません。
- どちらの helper も session lifecycle を持ちません。
- どちらの helper も callback / pending / foreground / normalized outcome の意味を変更しません。

維持した contract:

- emit 位置、trace 名、trace source、trace payload field set / order / names / values は維持します。
- `NewCaseVisibilityObservation.Complete(...)` 呼び出し位置は維持します。
- pending != completion、callback != completion、WindowActivate dispatch != completion、normalized outcome != completion、foreground outcome != completion は維持します。
- `case-display-completed` one-time emit、display session boundary、retry sequencing、foreground outcome semantics は維持します。

次 runtime 候補:

- 次に runtime を触る場合の候補は、R10/R11/R12 normalized outcome chain 呼び出し整理です。
- ただし、これはまだ GO ではありません。先に tests-first / safety net 評価が必要です。

現時点 STOP:

- display session lookup / one-time emit guard helper 化。
- callback raw facts adapter。
- foreground display-completable helper 化。
- route / dispatch shell 整理。
- R07/R09/R13/R14 をまたぐ runtime extraction。

### Phase 5 third runtime state

Phase 5 第三実装後の runtime refactor 完了点は次の 3 点です。

- completion hard gate yes/no decision の private helper 化。
- `case-display-completed` details payload assembly の private helper 化。
- R10/R11/R12 normalized outcome chain 呼び出しの private helper 化。

いずれも `TaskPaneRefreshOrchestrationService` 内の局所 helper 化です。owner 移動ではありません。

normalized outcome chain helper の意味:

- `CompleteNormalizedOutcomeChain(...)` は `CompleteVisibilityRecoveryOutcome(...)`、`CompleteRefreshSourceSelectionOutcome(...)`、`CompleteRebuildFallbackOutcome(...)` を既存順序で呼ぶだけです。
- helper は `TaskPaneRefreshAttemptResult` を受け取り、R10 -> R11 -> R12 の順に normalized outcome を付与した result を返すだけです。
- helper は completion 判定を持ちません。
- helper は foreground 判定を持ちません。
- helper は session lookup を持ちません。
- helper は one-time emit guard を持ちません。
- helper は `case-display-completed` emit を持ちません。
- helper は WindowActivate semantics を持ちません。
- helper は callback / pending の意味付けを持ちません。

未移動 owner:

- R13 foreground interpretation は未移動です。
- R14 completion gate は未移動です。
- `case-display-completed` emit owner は未移動です。
- display session boundary は未移動です。
- session lookup は未移動です。
- `IsCompleted` guard は未移動です。
- lock は未移動です。
- dictionary remove は未移動です。
- `NewCaseVisibilityObservation.Complete(...)` は未移動です。
- WindowActivate handling は未移動です。
- trace owner / payload contract は未移動です。

維持した contract:

- trace 名、trace source、trace payload field set / order / names / values は維持します。
- normal refresh path は R10/R11/R12 normalized outcome chain -> R13 foreground interpretation -> WindowActivate downstream observation -> R14 completion gate の距離を維持します。
- ready-show callback path は R10/R11/R12 normalized outcome chain -> R13 foreground interpretation -> R14 completion gate の距離を維持します。
- precondition skip path は R10/R11/R12 normalized outcome chain -> WindowActivate downstream observation -> return であり、R13/R14 へ進みません。
- normalized outcome != completion、foreground outcome != completion、callback != completion、pending != completion、WindowActivate dispatch != completion は維持します。
- `case-display-completed` one-time emit、display session boundary、retry sequencing、foreground outcome semantics は維持します。

次 runtime 候補:

- 第三実装後の次 runtime 候補はまだ GO ではありません。
- 次に runtime を触る場合は、改めて tests-first / safety net 評価を置きます。

現時点 STOP:

- foreground display-completable 判定 helper 化。
- display session lookup / one-time emit guard helper 化。
- callback raw facts adapter。
- route / dispatch shell 整理。
- R07/R09/R13/R14 横断 extraction。

### Phase 5 R13 trace details runtime state

Phase 5 R13 foreground trace details helper 化後の runtime refactor 完了点は次の 5 点です。

- completion hard gate yes/no decision の private helper 化。
- `case-display-completed` details payload assembly の private helper 化。
- R10/R11/R12 normalized outcome chain 呼び出しの private helper 化。
- R13 foreground execution result classification の private helper 化。
- R13 foreground trace details assembly の private helper 化。

いずれも `TaskPaneRefreshOrchestrationService` 内の局所 helper 化です。owner 移動ではありません。

R13 classification helper の意味:

- `ClassifyRequiredForegroundExecutionOutcome(...)` は execution result を foreground outcome に分類するだけです。
- `ExecutionAttempted && Recovered` の場合だけ `RequiredSucceeded` を返します。
- それ以外は現行通り `RequiredDegraded` を返します。
- `RequiredDegraded` を `RequiredFailed` へ丸めません。
- `RequiredDegraded` を success / failure / direct completion へ読み替えません。
- helper は foreground execution 呼び出し、trace emit、WindowActivate handling、R14 completion gate、`case-display-completed` emit、session lookup、one-time emit guard を持ちません。

R13 trace details helper の意味:

- `BuildForegroundRecoveryDecisionDetails(...)` は `foreground-recovery-decision` observation details の文字列 assembly だけを担います。field set / order は `reason -> foregroundRecoveryStarted -> foregroundSkipReason -> foregroundOutcomeStatus` です。
- `BuildFinalForegroundGuaranteeStartedDetails(...)` は `final-foreground-guarantee-started` observation details の文字列 assembly だけを担います。field set / order は `reason` です。
- `BuildFinalForegroundGuaranteeCompletedDetails(...)` は `final-foreground-guarantee-completed` observation details の文字列 assembly だけを担います。field set / order は `reason -> recovered -> foregroundOutcomeStatus` です。
- completed mapping は `recovered=true` の場合だけ `RequiredSucceeded`、`recovered=false` の場合は `RequiredDegraded` です。
- `RequiredDegraded` を `RequiredFailed`、success、direct completion へ丸めません。
- helper は foreground execution、WindowActivate handling、trace emit owner、R14 completion gate、`case-display-completed` emit、session lookup、one-time emit guard、callback / pending / normalized outcome の意味付けを持ちません。

未移動 owner:

- foreground execution 呼び出しは未移動です。
- trace action / source / emit position は未移動です。
- logger action / 発火順は未移動です。
- WindowActivate handling は未移動です。
- R14 completion gate は未移動です。
- `case-display-completed` emit owner は未移動です。
- display session boundary は未移動です。
- session lookup は未移動です。
- `IsCompleted` guard は未移動です。
- lock は未移動です。
- dictionary remove は未移動です。
- `NewCaseVisibilityObservation.Complete(...)` は未移動です。
- trace owner / payload contract は未移動です。

維持した freeze line:

- foreground outcome != completion。
- `RequiredDegraded` は success / failure / direct completion ではありません。
- `RequiredSucceeded` は input only です。
- `RequiredFailed` は completion gate を通しません。
- `NotRequired` は foreground success ではありません。
- `SkippedAlreadyVisible` は foreground success ではありません。
- callback != completion、pending != completion、WindowActivate dispatch != completion は維持します。
- `case-display-completed` one-time emit、display session boundary、trace contract、foreground outcome semantics は維持します。

次 runtime 候補:

- 次 runtime 候補はまだ GO ではありません。
- 次に runtime を触る場合は、foreground display-completable 判定 helper 化の tests-first 評価を先に置きます。

現時点 STOP:

- foreground display-completable helper 化。
- display session lookup / one-time emit guard helper 化。
- callback raw facts adapter。
- route / dispatch shell 整理。
- R07/R09/R13/R14 横断 extraction。

## 対象範囲

対象に含めるもの:

- created CASE 表示後の TaskPane 表示回復。
- `KernelCasePresentationService.ShowCreatedCase.PostRelease` 由来の ready-show。
- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` から入る window-dependent refresh 境界。
- visible pane already satisfied path と refresh path の収束。
- foreground guarantee と `case-display-completed` の成立条件。
- TaskPane 表示回復に隣接する suppression / protection。

対象に含めないもの:

- CASE 作成そのものの hidden create route 選択。
- hidden create session の workbook close / application quit / retained app cleanup。
- post-close white Excel prevention。
- 文書作成、会計書類セット作成、雛形登録の業務ルール。
- TaskPane の見た目、ボタン配置、タブ構成の仕様変更。

## 現行フロー要約

TaskPane 表示回復は、単一の直線フローではありません。現行では次の protocol unit が連鎖します。

1. `KernelCasePresentationService.OpenCreatedCase(...)` が created CASE の表示を開始する。
2. interactive route では `CaseWorkbookOpenStrategy.OpenHiddenForCaseDisplay(...)` が shared/current app 上で CASE workbook を hidden reopen する。
3. `KernelCasePresentationService.ShowCreatedCase(...)` が workbook window visibility ensure と without-showing recovery を実行する。
4. deferred presentation で transient suppression を release し、再度 visibility ensure してから CASE pane activation suppression を準備する。
5. `KernelCasePresentationService` が `ShowWorkbookTaskPaneWhenReady(...)` を呼び、ready-show request を出す。
6. `TaskPaneRefreshOrchestrationService` が ready-show を enqueue し、created-case display session を開始する。
7. `WorkbookTaskPaneReadyShowAttemptWorker` が ready-show attempt を実行する。
8. visible CASE pane が既に同じ workbook/window に表示済みなら already-visible path で成功相当にする。
9. already-visible でなければ `TaskPaneRefreshCoordinator` が refresh path を実行し、`TaskPaneManager` / `TaskPaneHostFlowService` が host reuse / render / show を行う。
10. `TaskPaneRefreshOrchestrationService` が visibility outcome、refresh source outcome、rebuild fallback outcome、foreground outcome を正規化する。
11. pane visible と foreground terminal が揃った場合だけ、同一 created-case display session に対して `case-display-completed` を 1 回だけ emit する。

重要な読み方:

- `pane visible`、`refresh completed`、`foreground guarantee completed`、`CASE display completed` は同義ではありません。
- `WindowActivate` は window-safe な refresh trigger であり、recovery owner、foreground owner、display completion owner ではありません。
- hidden-for-display / hidden create / retained hidden app-cache / post-close quit は隣接しますが、TaskPane 表示回復 owner へ昇格させません。

## Route

### created CASE ready-show route

入口:

- `KernelCasePresentationService.OpenCreatedCase(...)`
- `KernelCasePresentationService.ShowCreatedCase(...)`
- `KernelCasePresentationService.ExecuteDeferredPresentationEnhancements(...)`
- `TaskPaneRefreshOrchestrationService.ShowWorkbookTaskPaneWhenReady(...)`

route contract:

- interactive created CASE の post-release reason は `KernelCasePresentationService.ShowCreatedCase.PostRelease` として扱う。
- `display-handoff-completed` は ready-show request を `TaskPaneRefreshOrchestrationService` が受理した境界で記録する。
- `case-display-completed` は `TaskPaneRefreshOrchestrationService` だけが emit する。
- already-visible path でも refresh path でも、final completion owner は変えない。

### WorkbookActivate route

入口:

- `ThisAddIn.Application_WorkbookActivate(...)`
- `WorkbookLifecycleCoordinator.OnWorkbookActivate(...)`
- `TaskPaneRefreshOrchestrationService.TryRefreshTaskPane("WorkbookActivate", workbook, null)`

route contract:

- `WorkbookActivate` は workbook が active workbook として前面系の文脈に乗った後続イベントです。
- 入力 window は現行どおり `null` で入り、必要なら `TaskPaneRefreshCoordinator` 側で window resolve する。
- protection / suppression に当たる場合は refresh へ進めない。

### WindowActivate route

入口:

- `ThisAddIn.Application_WindowActivate(...)`
- `WindowActivatePaneHandlingService.Handle(...)`
- `ThisAddIn.RequestTaskPaneDisplayForTargetWindow(...)`
- `TaskPaneRefreshOrchestrationService.TryRefreshTaskPane(request, workbook, window)`

route contract:

- `WindowActivatePaneHandlingService` は `TaskPaneDisplayRequest.ForWindowActivate(...)` を作る。
- 分岐順は `case protection -> external workbook detection -> case pane suppression -> refresh dispatch` で固定する。
- `WindowActivateDispatchOutcome.Dispatched` は display success ではない。
- `WindowActivate` が refresh を trigger しても、`case-display-completed` は orchestration 側の completion 条件を満たした場合だけ成立する。

### WorkbookOpen route

入口:

- `ThisAddIn.Application_WorkbookOpen(...)`
- `WorkbookLifecycleCoordinator.OnWorkbookOpen(...)`

route contract:

- `WorkbookOpen` は workbook が開いた通知であり、window 安定境界ではありません。
- `WorkbookOpen` 直後に `workbook != null && window == null` の window-dependent refresh は `TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(...)` で skip する。
- `WorkbookOpen` 直後の `ActiveWorkbook` / `ActiveWindow` を前提に window resolve、表示、前面化、pane 対象決定を確定しない。

### retry fallback route

入口:

- ready-show attempts exhausted
- `TaskPaneRefreshOrchestrationService.ScheduleWorkbookTaskPaneRefresh(...)`
- `PendingPaneRefreshRetryService`

route contract:

- ready-show retry が尽きた場合だけ pending retry fallback へ handoff する。
- pending retry は workbook target を追い、対象 workbook を見失っても active CASE context が残る場合は active refresh fallback を継続する。
- pending retry は window resolve や refresh dispatch の意味を変えるものではなく、fallback scheduling の owner です。

### R07 pending fallback handoff current-state

`TaskPaneRefreshOrchestrationService.ScheduleWorkbookTaskPaneRefresh(...)` は、単なる delayed timer schedule helper ではありません。現行では次を 1 つの protocol entry として束ねています。

- `ready-show-fallback-handoff` / `wait-ready-fallback-handoff` の trace。
- `WorkbookOpen` window-dependent skip。
- workbook target tracking。
- workbook から pane 対象 window への resolve。
- pending timer 開始前の immediate refresh。
- pending retry を開始するかどうかの decision。

この entry は、少なくとも次の 2 つの caller / reason から使われます。

- created CASE ready-show exhaustion 後の handoff。代表 reason は `KernelCasePresentationService.ShowCreatedCase.PostRelease`。
- `KernelHomeForm.OpenSheet.PostClose` の workbook-target delayed refresh entry。

したがって、現時点では `ScheduleWorkbookTaskPaneRefresh(...)` を ready-show exhaustion 専用 entry として runtime extraction しません。ready-show handoff と workbook-target delayed refresh entry の二重性を持つ orchestration boundary として扱います。

immediate refresh は pending timer 開始前の refresh re-entry です。immediate refresh が success した場合でも、それだけでは recovered、display recovery completed、`case-display-completed` のいずれも意味しません。completion は existing orchestration completion chain の条件を満たした場合だけ成立します。

`WorkbookOpen` skip は null guard ではなく window stability boundary の runtime stabilization contract です。`ready-show-fallback-handoff` trace 後であっても、`reason == "WorkbookOpen" && workbook != null && window == null` の場合は pending retry start へ進めず、後続の `WorkbookActivate` / `WindowActivate` 側へ委ねます。

`PendingPaneRefreshRetryService` 内の active CASE context fallback は、tracked workbook を見失った時の target-lost resiliency fallback です。これは completion fallback ではなく、成功時も refresh / outcome / completion chain に戻れた場合だけ display completion の材料になります。

### R08 active CASE fallback current-state

`PendingPaneRefreshRetryService` の pending retry tick は、tracked workbook を優先し、tracked workbook を見失った場合だけ active CASE context fallback を試します。

現行分岐:

- attempts が残っていない場合は timer を停止します。
- tracked workbook が見つかる場合は、その workbook の window resolve を行い、workbook target refresh を試します。
- tracked workbook を見失ったが active context が CASE の場合は、active CASE context fallback として `TryRefreshTaskPane(reason, null, null)` を試します。
- tracked workbook を見失い、active context が null または CASE 以外の場合は timer を停止します。
- tracked workbook route または active CASE fallback route の refresh success 時は timer を停止します。

active CASE fallback は、tracked workbook を見失った場合でも active workbook が CASE として解決できるなら refresh attempt を継続する target-lost resiliency fallback です。これは completion fallback、foreground fallback、display session completion、created CASE display completion の代替経路ではありません。

truth table:

| tracked workbook exists | active context is CASE | attempts remaining | refresh attempted | refresh target | timer continues | completion meaning | trace / outcome meaning |
| --- | --- | --- | --- | --- | --- | --- | --- |
| true | 該当なし | yes | yes | tracked workbook + resolved pane window | refresh success なら stop。refresh failure かつ attempts が残る場合だけ継続。 | active fallback 自体は completion を emit しない。tracked refresh success も `case-display-completed` ではない。 | `defer-retry-start` / `defer-retry-end` は workbook-target retry attempt の観測。`refreshed=true` は refresh attempt result であり completion trace ではない。 |
| false | true | yes | yes | active CASE context via `TryRefreshTaskPane(reason, null, null)` | refresh success なら stop。refresh failure かつ attempts が残る場合だけ継続。 | active fallback 自体は completion を emit しない。fallback refresh success も completion fallback ではない。 | `defer-active-context-fallback-start` / `defer-active-context-fallback-end` は target-lost resiliency fallback の観測。`refreshed=true` は active refresh attempt result。 |
| false | false | yes | no | none | stop | active fallback 自体は completion を emit しない。stop は completion ではない。 | `defer-active-context-fallback-stop` は active CASE fallback 不成立の観測。success / recovered / foreground を意味しない。 |
| any | any | no | no | none | stop | active fallback 自体は completion を emit しない。attempts exhausted は completion ではない。 | attempts exhausted による timer stop は retry lifecycle の観測であり、display failure / completion trace ではない。 |

stop conditions:

- refresh success。
- attempts exhausted。
- active context が CASE でない。
- tracked workbook / active fallback のどちらでも refresh attempt できない。

ただし、stop は completion ではありません。pending retry が stopped になっても、`case-display-completed` は emit されません。created CASE display completion は、既存の display session boundary と normalized outcome chain が pane visible、visibility terminal / display-completable、foreground terminal / display-completable を満たした場合だけ成立します。

trace / outcome の読み方:

- `defer-retry-start` / `defer-retry-end` / `defer-active-context-fallback-start` / `defer-active-context-fallback-end` / `defer-active-context-fallback-stop` は観測 trace です。
- trace source string、trace payload、trace 名は現行契約として維持します。
- `refreshed=true` は refresh attempt result であり、recovered event、foreground success、display session completion、`case-display-completed` のいずれでもありません。
- active CASE fallback から戻った attempt result も、orchestration 側の completion owner が既存条件を満たすまで completion にはなりません。

## Retry

現行実装値:

| retry | owner | 値 | 意味 |
| --- | --- | --- | --- |
| ready-show max attempts | `WorkbookTaskPaneReadyShowAttemptWorker` / `TaskPaneDisplayRetryCoordinator` | `2` | ready-show の即時 attempt と delayed attempt の上限。 |
| ready-show retry delay | `TaskPaneReadyShowRetryScheduler` | `80ms` | ready-show attempt 2 を発火させる delay。 |
| pending retry interval | `PendingPaneRefreshRetryService` | `400ms` | ready-show exhaustion 後の fallback refresh interval。 |
| pending retry max attempts | `PendingPaneRefreshRetryService` | `3` | fallback refresh の attempts。 |
| suppression / protection duration | `KernelHomeCasePaneSuppressionCoordinator` | `5秒` | activation refresh suppression と foreground protection の duration。 |

不明:

- これらの数値が正式な業務仕様値か、現行実装上の経験値かは docs / code だけでは確定しません。

変更禁止:

- ready-show `2` attempts と `80ms` retry の順序を変えない。
- ready-show exhaustion 前に pending retry `400ms` route へ落とさない。
- pending retry `3` attempts と active CASE context fallback を削らない。
- retry 失敗を display success に丸めない。

## Ready-Show

ready-show は CASE 表示直後の TaskPane 表示を安定させる delayed display route です。

現行順序:

1. `KernelCasePresentationService` が transient suppression を release する。
2. `WorkbookWindowVisibilityService.EnsureVisible(...)` を再実行する。
3. `SuppressUpcomingCasePaneActivationRefresh(...)` を設定する。
4. `ShowWorkbookTaskPaneWhenReady(...)` を呼ぶ。
5. `TaskPaneRefreshOrchestrationService` が `ready-show-enqueued` を記録する。
6. `TaskPaneRefreshOrchestrationService` が created-case display session を開始し、`created-case-display-session-started` と `display-handoff-completed` を記録する。
7. `WorkbookTaskPaneReadyShowAttemptWorker` が attempt 1 を実行する。
8. attempt 1 だけ `WorkbookWindowVisibilityService.EnsureVisible(...)` を実行する。
9. `ResolveWorkbookPaneWindow(..., activateWorkbook: true)` で window を解決する。
10. visible CASE pane が既に対象 workbook/window にある場合は `taskpane-already-visible` で refresh を skip する。
11. そうでなければ `TryRefreshTaskPane(...)` へ handoff する。
12. attempt が成功しなければ `80ms` 後に attempt 2 を schedule する。
13. attempt 2 も成功しなければ `ready-show-fallback-handoff` で pending retry route へ渡す。

変更禁止:

- `ReleaseWorkbook -> EnsureVisible -> SuppressUpcomingCasePaneActivationRefresh -> ShowWorkbookTaskPaneWhenReady` の順を変えない。
- attempt 1 だけの pre-visibility ensure を、無根拠に全 attempt へ広げない。
- visible CASE pane early-complete を accounting などへ広げない。
- already-visible path で `case-display-completed` を worker 側に戻さない。

## Foreground Outcome

foreground guarantee は、pane visible 後に foreground obligation が残るかどうかを terminal outcome にする protocol です。

現行 owner:

- decision / normalized outcome / trace owner: `TaskPaneRefreshOrchestrationService`
- execution bridge: `TaskPaneRefreshCoordinator.ExecuteFinalForegroundGuaranteeRecovery(...)`
- execution primitive: `ExcelWindowRecoveryService`

現行 decision:

- `ForegroundGuaranteeOutcome.SkippedAlreadyVisible`
  - visible pane already satisfied path など、既に visible と判断された場合。
- `ForegroundGuaranteeOutcome.NotRequired`
  - refresh が成功していない、pane visible ではない、refresh completed ではない、window がない、recovery service がない場合など。
- `ForegroundGuaranteeOutcome.RequiredSucceeded`
  - foreground recovery execution が attempted かつ recovered の場合。
- `ForegroundGuaranteeOutcome.RequiredDegraded`
  - execution は attempted したが recovery が false の場合。
- `ForegroundGuaranteeOutcome.Unknown`
  - execution pending など、途中状態の normalized value。

現行 completion 条件:

- `case-display-completed` は `attemptResult.IsRefreshSucceeded`、`attemptResult.IsPaneVisible`、`VisibilityRecoveryOutcome.IsTerminal`、`VisibilityRecoveryOutcome.IsDisplayCompletable`、`ForegroundGuaranteeOutcome.IsTerminal`、`ForegroundGuaranteeOutcome.IsDisplayCompletable` が揃った場合だけ成立する。
- `refresh completed` は補助条件であり、already-visible path では必須ではない。

### Foreground outcome decision contract current-state

foreground outcome chain は、window resolve、refresh attempt、visibility recovery、display session completion とは別の decision contract です。現行 code では completion emit と近い owner にありますが、foreground decision step 自体は `case-display-completed` emit owner ではありません。

別軸として読む fact:

| fact axis | current owner / source | 意味 | 読み替え禁止 |
| --- | --- | --- | --- |
| window resolve | R09 `ResolveWorkbookPaneWindow(...)` / refresh result の `ForegroundWindow` | foreground decision に使える window fact があるか。 | resolved window = foreground success / completion ではない。 |
| refresh / pane visibility | refresh attempt result と visibility outcome | pane visible と visibility recovery の terminal / display-completable 条件。 | visible / refresh success を foreground success へ丸めない。 |
| foreground required decision | `CompleteForegroundGuaranteeOutcome(...)` | foreground execution が必要条件を満たすか。 | required / not-required を visibility outcome と同じ enum 意味にしない。 |
| foreground execution result | `TaskPaneRefreshCoordinator.ExecuteFinalForegroundGuaranteeRecovery(...)` | required path の raw recovered fact。 | execution result 単体を `case-display-completed` と読まない。 |
| display session completion | `TryCompleteCreatedCaseDisplaySession(...)` | created CASE display session の one-time emit 条件。 | foreground outcome owner へ emit 責務を移さない。 |

foreground outcome の主要意味:

| outcome | current-state meaning | completion との距離 |
| --- | --- | --- |
| `RequiredSucceeded` | foreground guarantee が必要で、execution が attempted かつ recovered。 | display-completable terminal だが、それ単体では `case-display-completed` ではない。 |
| `RequiredDegraded` | foreground guarantee が必要で execution は attempted したが、success guarantee まではできなかった。 | display-completable terminal として扱うが success へ読み替えない。failure へも雑に丸めない。 |
| `NotRequired` | foreground guarantee の required 条件が揃わない、または foreground path 非該当。 | foreground success ではなく、他条件が揃った場合に completion input になれる terminal fact。 |

foreground outcome は completion input にはなり得ます。ただし `case-display-completed` の one-time emit は、created CASE display session、pane visible、visibility terminal / display-completable、foreground terminal / display-completable が揃った場合だけ orchestration 側の existing completion chain が行います。

WindowActivate との距離:

- `WindowActivate` dispatch / downstream observation は route observation であり、foreground success でも completion でもありません。
- `WindowActivateDownstreamObservation` が foreground status を trace payload に含める場合でも、その trace は `case-display-completed` emit owner ではありません。

変更禁止:

- foreground decision owner を `TaskPaneRefreshCoordinator` や `ExcelWindowRecoveryService` へ戻さない。
- `WindowActivate` 発火だけを foreground terminal とみなさない。
- foreground recovery の実行条件を広げない。
- `RequiredDegraded` を勝手に failure / success の別意味へ読み替えない。

## Window Resolve

window resolve は `TaskPaneRefreshOrchestrationService.WorkbookPaneWindowResolver.Resolve(...)` が持つ current-state helper です。

R09 current-state としては、`ResolveWorkbookPaneWindow(...)` は単なる null guard ではありません。TaskPane display recovery において、workbook は特定できているが pane の対象 window がまだ使えるかどうかを確定する `window availability boundary` です。

ただし R09 は、次の owner ではありません。

- completion 判定 owner ではありません。
- foreground decision owner ではありません。
- retry success owner ではありません。
- fallback start / stop owner ではありません。

R09 が返すのは、現時点で pane refresh / visible-pane 判定 / context 解決に渡せる `Excel.Window` があるかどうかだけです。`window != null` は display completed ではなく、`window == null` は retry 開始や fallback 成功を意味しません。未解決時の扱いは caller route の protocol に委ねます。

成功条件:

- 対象 workbook の first visible window を取得できる。
- または active workbook が対象 workbook と一致し、active window を取得できる。

attempt:

- `WorkbookPaneWindowResolveAttempts = 2`
- `WorkbookPaneWindowResolveDelayMs = 80`
- `activateWorkbook=true` の場合だけ `ExcelInteropService.ActivateWorkbook(workbook)` を呼ぶ。

trace:

- `resolve-window-start`
- `resolve-window-state`
- `resolve-window-success`
- `resolve-window-success-active-window`
- `resolve-window-retry`
- `resolve-window-failed`

変更禁止:

- `WorkbookOpen` 直後の window 未解決を推測で補わない。
- `WindowActivate` route で event window が渡っている場合に、event trigger と activation primitive を同一 owner とみなさない。
- context-less workbook 推測や暗黙 open を window resolve の fallback にしない。

### R09 activateWorkbook route matrix

`activateWorkbook` は route 固有の activation request です。foreground guarantee の成功、display completion、retry success とは同一視しません。

| route | R09 呼び出し / window source | `activateWorkbook` | window unresolved の意味 | retry / fallback への委譲 | completion への接続 | foreground semantics |
| --- | --- | --- | --- | --- | --- | --- |
| ready-show immediate / wait-ready path | `WorkbookTaskPaneReadyShowAttemptWorker.TryShowWorkbookTaskPaneOnce(...)` が `ResolveWorkbookPaneWindow(...)` を呼ぶ | `true` | ready-show attempt の pane 対象 window がまだ unavailable。attempt result は refresh success として確定しない。 | attempt 1 失敗なら `80ms` attempt 2、attempt 2 も失敗した場合だけ R07 pending fallback handoff。 | worker では completion しない。orchestration の callback が visibility / source / rebuild / foreground outcome を正規化した後だけ `case-display-completed` 候補になる。 | activation request はするが foreground guarantee success ではない。 |
| R07 `ScheduleWorkbookTaskPaneRefresh(...)` immediate refresh | pending timer 開始前に `ResolveWorkbookPaneWindow(...)` で workbook target window を準備 | `false` | immediate refresh に渡す window がないだけ。`WorkbookOpen` skip 以外では、refresh path 側が再度 fail-closed / context 解決する。 | immediate refresh が success しない場合だけ pending retry `400ms / 3 attempts` へ進む。 | immediate refresh success は completion ではない。created CASE reason かつ existing completion chain 条件を満たした場合だけ completion。 | foreground semantics は持たない。 |
| R08 pending retry tick | tracked workbook を見つけた tick で `ResolveWorkbookPaneWindow(...)` を呼ぶ | `true` | target workbook の window availability がまだ不十分。active CASE context fallback に移れる場合がある。 | tick の refresh attempt が success しなければ attempts を消費して継続。target を見失った場合は active CASE context fallback。 | retry tick success は completion ではない。refresh / outcome / display session boundary に戻れた場合だけ completion 候補。 | activation request は window resolve / refresh attempt のためで、foreground guarantee success ではない。 |
| `WorkbookOpen` skip / stabilization boundary | `TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(...)` が R09 呼び出し前に止める | 該当なし | workbook はあるが window-dependent refresh の安定境界に達していない。 | `ScheduleWorkbookTaskPaneRefresh(...)` では pending retry start へ進まず、後続 `WorkbookActivate` / `WindowActivate` に委ねる。 | skip は success / completion / fallback start ではない。 | foreground semantics は持たない。 |
| `WindowActivate` downstream observation | event window が `TaskPaneDisplayRequest` とともに渡る。event window がある場合、R09 resolve は不要 | 通常は該当なし。event window が欠け coordinator が補う場合も `false` | event / downstream refresh の観測であり、display success ではない。 | WindowActivate dispatch 自体は pending fallback を開始しない。 | `WindowActivateDispatchOutcome.Dispatched` や downstream outcome trace は completion ではない。 | WindowActivate は foreground owner ではない。 |
| `KernelHomeForm.OpenSheet.PostClose` delayed refresh | `KernelHomeForm` が `ScheduleWorkbookTaskPaneRefresh(displayedWorkbook, "KernelHomeForm.OpenSheet.PostClose")` に入る | immediate prepare は `false`。pending retry tick は `true` | HOME post-close delayed refresh の workbook target window がまだ unavailable。created CASE display reason ではない。 | immediate refresh が success しなければ pending retry に委ねる。 | created CASE reason ではないため、`case-display-completed` emit route ではない。 | foreground semantics は持たない。 |
| foreground guarantee path | R09 ではなく、refresh result の `ForegroundWindow` と foreground recovery service の availability を使う | 該当なし | foreground decision に必要な window fact が欠ける。 | retry / fallback owner ではない。 | `ForegroundGuaranteeOutcome` が terminal / display-completable になった場合だけ completion 条件の一部になる。 | R13 / Phase 5 候補。window resolve、activation request、foreground outcome、display-completable terminal を同一視しない。 |

### WorkbookOpen stabilization boundary

`WorkbookOpen` は window 安定境界ではありません。`WorkbookOpen + workbook exists + window unresolved` の場合、window-dependent refresh を即時実行しません。

- `WorkbookOpen` 直後の skip は incidental null guard ではなく、runtime stabilization contract です。
- skip は後続イベント、`WindowActivate`、または route が許す pending retry へ委ねる境界です。
- `ScheduleWorkbookTaskPaneRefresh(...)` では、`WorkbookOpen` skip に当たると pending retry start へ進みません。
- skip を success / completion / fallback start と読みません。

### WindowActivate / pending retry / foreground との距離

- `WindowActivate` downstream observation は window availability / route observation の一部です。`case-display-completed` emit owner ではありません。
- R08 pending retry tick の `activateWorkbook=true` は window resolve / refresh attempt のための request です。foreground guarantee success や completion を意味しません。
- foreground guarantee path は R13 / Phase 5 候補です。R09 docs では `window resolve`、`activation request`、`foreground outcome`、`display-completable terminal` を分けて扱います。

## Trace

trace 名は観測契約です。現行名を変更しません。

created CASE display:

- `case-workbook-open-started`
- `case-workbook-open-completed`
- `initial-recovery-completed`
- `post-release-suppression-prepared`
- `ready-show-requested`
- `ready-show-enqueued`
- `created-case-display-session-started`
- `display-handoff-completed`
- `case-display-completed`

ready-show:

- `ready-show-attempt`
- `ready-show-attempt-result`
- `taskpane-already-visible`
- `ready-show-attempts-exhausted`
- `ready-show-fallback-handoff`

refresh / visibility / foreground:

- `taskpane-refresh-started`
- `taskpane-refresh-completed`
- `visibility-recovery-decision`
- `visibility-recovery-{status}`
- `foreground-recovery-decision`
- `final-foreground-guarantee-started`
- `final-foreground-guarantee-completed`

refresh source / rebuild fallback:

- `refresh-source-selected`
- `refresh-source-degraded`
- `refresh-source-fallback`
- `refresh-source-rebuild-required`
- `refresh-source-failed`
- `refresh-source-not-reached`
- `refresh-source-unknown`
- `rebuild-fallback-required`
- `rebuild-fallback-{status}`

WindowActivate:

- `display-refresh-trigger-observed`
- `display-refresh-trigger-ignored`
- `display-refresh-trigger-deferred`
- `display-refresh-trigger-dispatched`
- `display-refresh-trigger-failed`
- `window-activate-display-refresh-trigger-start`
- `window-activate-display-refresh-trigger-outcome`

protection / suppression:

- `WorkbookActivateProtection action=evaluate`
- `WorkbookActivateProtection action=start`
- `WorkbookActivateProtection action=ignore`
- `WindowActivateProtection action=ignore`
- `TaskPaneRefreshProtection action=ignore`
- `Case pane activation suppression prepared`
- `Case pane refresh suppressed`
- `Case pane activation suppression cleared`

## Fail-Closed 条件

TaskPane 表示回復では、欠けた事実を推測で補完しません。

固定する fail-closed 条件:

- `KernelCaseCreationResult.Success == false` の場合、created CASE 表示へ進まない。
- CASE workbook path が空なら表示へ進めない。
- opened workbook が null なら表示へ進めない。
- `WorkbookOpen` 直後に `workbook != null && window == null` なら window-dependent refresh を skip する。
- protection 中は `WorkbookActivate` / `WindowActivate` / `TaskPaneRefresh` を無理に通さない。
- case pane activation suppression 中は refresh dispatch へ進めない。
- `_workbookSessionService` や `_taskPaneManager` が欠ける場合は refresh を skip する。
- taskPane refresh suppression count が 0 より大きい場合は refresh を skip する。
- context が受理できない場合は、必要に応じて対象 window の pane を hide し、refresh success にはしない。
- unknown role / missing window key は host-flow precondition で hide / skip する。
- `case-display-completed` は pane visible と visibility / foreground の terminal display-completable outcome が揃うまで emit しない。
- `WindowActivateDispatchOutcome.Dispatched` を display completion とみなさない。
- hidden session cleanup、white Excel prevention、foreground recovery を相互の代替 owner として扱わない。

不明:

- foreground が degraded した場合の user-facing guidance は未定義です。
- すべての環境で `WorkbookActivate` と `WindowActivate` のどちらを最終安全境界とみなすべきかは未確定です。

## Window-Dependent Refresh Policy

正本:

- `TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(...)`

現行条件:

```csharp
reason == "WorkbookOpen" && workbook != null && window == null
```

policy の意味:

- pure 判定です。
- ログ出力、状態変更、COM メンバーアクセス、UI 操作を持ちません。
- `TaskPaneRefreshOrchestrationService` と `TaskPaneRefreshCoordinator` はこの policy の利用者であり、同じ skip 条件を個別に持ちません。

変更禁止:

- `WorkbookOpen` を window 安定境界へ昇格させない。
- `ActiveWorkbook` / `ActiveWindow` の一時状態を根拠に表示対象を確定しない。
- policy に COM access、状態変更、ログ出力、UI 操作を追加しない。

## Owner

current owner:

| 領域 | current owner | 読み方 |
| --- | --- | --- |
| created CASE presentation | `KernelCasePresentationService` | wait UI、known path、transient suppression、initial recovery、ready-show request、cursor、one-shot promotion が集中している。 |
| hidden-for-display open | `CaseWorkbookOpenStrategy` | shared/current app 上の hidden reopen と previous window restore。TaskPane completion owner ではない。 |
| ready-show enqueue / display session / final completion | `TaskPaneRefreshOrchestrationService` | request acceptance、retry/fallback、normalized outcomes、`case-display-completed` を束ねる。 |
| ready-show attempt | `WorkbookTaskPaneReadyShowAttemptWorker` | 1 attempt の visibility ensure、window resolve、already-visible 判定、refresh delegate 呼び出し。 |
| pending retry | `PendingPaneRefreshRetryService` | `400ms` fallback retry、workbook target / active CASE fallback の状態管理。 |
| refresh execution | `TaskPaneRefreshCoordinator` | pre-context recovery、window resolve、context resolve、TaskPaneManager 呼び出し、raw result 返却。 |
| pane visible state | `TaskPaneHostFlowService` / `TaskPaneDisplayCoordinator` | host reuse / render / show と visible pane 判定。 |
| workbook visibility primitive | `WorkbookWindowVisibilityService` | workbook window visible ensure。foreground decision は持たない。 |
| full recovery primitive | `ExcelWindowRecoveryService` | ScreenUpdating、window recovery、activation、foreground promotion の実行 owner。 |
| WorkbookActivate entry | `WorkbookLifecycleCoordinator` | workbook activation event と refresh dispatch。 |
| WindowActivate dispatch | `WindowActivatePaneHandlingService` | request 化、protection、external detection、suppression、dispatch。 |
| suppression / protection state | `KernelHomeCasePaneSuppressionCoordinator` | Kernel HOME suppression、CASE pane activation suppression、foreground protection が同居している。 |
| VSTO boundary / adapter | `ThisAddIn` | event wiring、display entry bridge、TaskPane create/remove adapter が残る。 |

## 呼び出し順序

created CASE 表示回復:

1. `OpenCreatedCase(...)`
2. `RegisterKnownCasePath(...)`
3. `TransientPaneSuppressionService.SuppressPath(...)`
4. `OpenCreatedCaseWorkbook(...)`
5. `OpenHiddenForCaseDisplay(...)` または `OpenVisibleWorkbook(...)`
6. `case-workbook-open-completed`
7. `EnsureWorkbookWindowVisibleBeforeInitialRecovery(...)`
8. `ExcelWindowRecoveryService.TryRecoverWorkbookWindowWithoutShowing(..., bringToFront: false)`
9. `initial-recovery-completed`
10. `TransientPaneSuppressionService.ReleaseWorkbook(...)`
11. `EnsureWorkbookWindowVisibleBeforeReadyShow(...)`
12. `SuppressUpcomingCasePaneActivationRefresh(...)`
13. `ShowWorkbookTaskPaneWhenReady(...)`
14. `ready-show-enqueued`
15. `created-case-display-session-started`
16. `display-handoff-completed`
17. ready-show attempt
18. visible already satisfied path または refresh path
19. foreground guarantee outcome
20. `case-display-completed`

refresh path:

1. `ThisAddIn.RefreshTaskPane(...)`
2. `TaskPaneRefreshOrchestrationService.TryRefreshTaskPane(...)`
3. `TaskPaneRefreshPreconditionPolicy.DecideRefreshPrecondition(...)`
4. `TaskPaneRefreshCoordinator.TryRefreshTaskPane(...)`
5. dependency / suppression check
6. WorkbookOpen window-dependent skip check
7. pre-context recovery
8. pane window resolve
9. `WorkbookSessionService.ResolveContext(...)`
10. context accept / reject
11. `TaskPaneManager.RefreshPaneWithOutcome(...)`
12. `TaskPaneHostFlowService` reuse / render / show
13. visibility / refresh source / rebuild fallback outcome normalization
14. foreground guarantee decision / execution
15. created-case display completion check

WindowActivate route:

1. `ThisAddIn.Application_WindowActivate(...)`
2. `WorkbookEventCoordinator.OnWindowActivate(...)`
3. `ThisAddIn.HandleWindowActivateEvent(...)`
4. `WindowActivatePaneHandlingService.Handle(...)`
5. `TaskPaneDisplayRequest.ForWindowActivate(...)`
6. protection check
7. external workbook detection
8. case pane suppression check
9. `ThisAddIn.RequestTaskPaneDisplayForTargetWindow(...)`
10. `PaneDisplayPolicy.Decide(...)`
11. `RefreshTaskPane(...)`
12. `TaskPaneRefreshOrchestrationService.TryRefreshTaskPane(...)`

## 変更禁止順序

絶対に崩さない順序:

- `WorkbookOpen -> WorkbookActivate -> WindowActivate` を、window stability の読みにおいて混同しない。
- created CASE deferred presentation は `ReleaseWorkbook -> EnsureVisible -> SuppressUpcomingCasePaneActivationRefresh -> ShowWorkbookTaskPaneWhenReady` の順を維持する。
- ready-show は `attempt 1 -> 80ms retry attempt 2 -> pending retry fallback` の順を維持する。
- `case protection -> external workbook detection -> case pane suppression -> refresh dispatch` の WindowActivate gate 順を維持する。
- refresh path は `precondition -> pre-context recovery -> window resolve -> context resolve -> context accept -> host flow -> normalized outcomes -> foreground guarantee -> completion check` の順を維持する。
- foreground guarantee は refresh result / pane visible の後にだけ判定する。
- post foreground protection は foreground execution bridge 後に開始する。
- `case-display-completed` は lower-level worker / coordinator / host-flow から emit しない。
- hidden session cleanup や post-close quit と TaskPane display completion を混ぜない。

変更禁止の理由:

- 順序自体が flicker、white Excel、reentrant activation、pane 二重生成、window unresolved refresh を避ける安全装置になっているため。
- この順序を固定しておくことで、Phase 4 以降に owner だけを切り出しても挙動差分を検出しやすくなるため。

## 現在の Orchestration 集中箇所

### `TaskPaneRefreshOrchestrationService`

集中している責務:

- normal refresh entry。
- ready-show enqueue。
- ready-show retry timer。
- pending retry fallback 生成。
- window resolver helper。
- refresh precondition policy decision 呼び出し。
- created-case display session state。
- visibility outcome 正規化。
- refresh source outcome 正規化。
- rebuild fallback outcome 正規化。
- foreground guarantee decision / trace。
- `case-display-completed` emit。

現在のリスク:

- retry、trace、completion、foreground、window resolve が近接しており、1 変更が表示完了条件へ波及しやすい。
- ただし現時点では final completion owner がここに寄っているため、いきなり細かく剥がすと completion 条件が散る。

### `KernelCasePresentationService`

集中している責務:

- created CASE wait UI。
- known path registration。
- transient suppression。
- hidden / visible open route selection。
- initial visibility ensure。
- without-showing recovery。
- post-release suppression setup。
- ready-show request。
- cursor positioning。
- wait UI close。
- `NewCaseDefault` 以外の one-shot foreground promotion。

現在のリスク:

- CASE 表示の presentation owner と TaskPane display protocol request owner が近接している。
- hidden-for-display open、TaskPane ready-show、cursor 移動、foreground promotion が 1 メソッド列に同居している。

### `ThisAddIn`

集中している責務:

- Excel event subscription / handler。
- `RefreshTaskPane(...)` bridge。
- `ShowWorkbookTaskPaneWhenReady(...)` bridge。
- `RequestTaskPaneDisplayForTargetWindow(...)` display entry。
- VSTO `CustomTaskPane` create/remove adapter。
- protection / suppression bridge。
- retained hidden app-cache shutdown call。

現在のリスク:

- VSTO adapter、event boundary、display request bridge が同じ add-in class に残っている。
- `WindowActivate` request source と downstream string reason の変換境界が add-in 付近にある。

### `KernelHomeCasePaneSuppressionCoordinator`

集中している責務:

- Kernel HOME display suppression。
- CASE pane activation suppression。
- CASE foreground protection。

現在のリスク:

- suppression と protection はどちらも activation / refresh を止めるが、意味が異なる。
- `WorkbookActivate` / `WindowActivate` / `TaskPaneRefresh` の 3 入口にまたがるため、owner を分ける前に trace / order 固定が必要。

### `TaskPaneManager` 周辺

集中している責務:

- facade entry。
- host registry / lifecycle との接続。
- render / show / reuse flow。
- visible CASE pane 判定。

現在のリスク:

- host visible state と display completion が近接して見える。
- ただし `case-display-completed` owner ではないため、ここを completion owner に戻さない。

## 本来あるべき責務境界

この節は service 増加を目的にしません。変更理由ごとの境界を固定するための target boundary です。

| 変更理由 | 本来の境界 | 現在の主な owner | 将来減らせる変更コスト |
| --- | --- | --- | --- |
| created CASE presentation preparation | wait UI、suppression、initial recovery、ready-show request まで | `KernelCasePresentationService` | CASE 表示前処理と TaskPane display protocol の混同を減らす。 |
| display protocol session | ready-show acceptance、attempt outcome、fallback、foreground outcome、final completion | `TaskPaneRefreshOrchestrationService` | `case-display-completed` の成立条件を 1 箇所で守る。 |
| ready-show attempt | 1 attempt の window resolve、already-visible 判定、refresh delegate 呼び出し | `WorkbookTaskPaneReadyShowAttemptWorker` | retry / attempt 本体の変更を completion owner から分ける。 |
| retry state | ready retry scheduler と pending retry state | `TaskPaneReadyShowRetryScheduler` / `PendingPaneRefreshRetryService` | retry 値や fallback 監視の変更理由を局所化する。 |
| pane visible state | host reuse、render、show、visible metadata | `TaskPaneHostFlowService` / `TaskPaneDisplayCoordinator` | host 表示の変更を display completion から分ける。 |
| visibility primitive | lightweight visible ensure と full recovery primitive | `WorkbookWindowVisibilityService` / `ExcelWindowRecoveryService` | window mutation 条件を display protocol から分ける。 |
| foreground guarantee | decision / outcome / trace と execution primitive の分離 | `TaskPaneRefreshOrchestrationService` / `TaskPaneRefreshCoordinator` / `ExcelWindowRecoveryService` | foreground 条件変更と Win32/COM primitive 変更を分ける。 |
| event trigger | WorkbookOpen / WorkbookActivate / WindowActivate の capture と dispatch | `WorkbookLifecycleCoordinator` / `WindowActivatePaneHandlingService` / `ThisAddIn` | event 境界変更を display completion 条件から分ける。 |
| suppression / protection state | CASE pane activation suppression と foreground protection | `KernelHomeCasePaneSuppressionCoordinator` | 「止める理由」単位で将来切り出せる。 |
| VSTO adapter | `CustomTaskPane` create/remove | `ThisAddIn` | Office adapter 変更を domain/display protocol から分ける。 |

## Future Change Cost を下げる観点

この docs 固定により、次の変更コストを下げる。

- ready-show retry を観測・調整したい場合、attempt owner / retry owner / completion owner を混同しなくて済む。
- foreground 表示回復を調整したい場合、decision owner と execution primitive owner を分けて見られる。
- `WindowActivate` 周辺を直したい場合、event trigger と display completion を混同しなくて済む。
- TaskPane host reuse を整理したい場合、pane visible state と CASE display completed の違いを守れる。
- hidden Excel / white Excel の lifecycle 修正と TaskPane 表示回復を同じ変更に混ぜにくくなる。

## 今回行わないこと

- コード変更。
- service / helper 抽出。
- retry 値変更。
- trace 名変更。
- route contract 変更。
- fail-closed 条件変更。
- foreground / visibility recovery 条件変更。
- `WorkbookOpen` window-dependent skip 条件変更。
- hidden create route / hidden-for-display route / retained hidden app-cache cleanup 変更。
- post-close white Excel prevention 変更。
- build / test / `DeployDebugAddIn` 実行。

## 次フェーズへの渡し方

Phase 1 では、この文書をもとに responsibility inventory を作る。最初の inventory 候補は次の単位にする。

1. `TaskPaneRefreshOrchestrationService` の display session / retry / outcome / window resolve / trace responsibility inventory。
2. `KernelCasePresentationService` の presentation preparation / ready-show request / cursor / foreground adjacency inventory。
3. `WindowActivatePaneHandlingService` と `ThisAddIn` の event trigger / display entry inventory。
4. `KernelHomeCasePaneSuppressionCoordinator` の suppression / protection state inventory。

Phase 4 で ownership 分離に進む場合も、この文書の route、trace、retry、fail-closed、変更禁止順序を先に固定線として扱う。Phase 3 の変更禁止契約は `docs/taskpane-display-recovery-freeze-line.md` を正本とする。
