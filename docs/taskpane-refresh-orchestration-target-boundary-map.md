# TaskPaneRefreshOrchestrationService Target Boundary Map

## 位置づけ

この文書は、TaskPane 表示回復領域の Phase 2「理想責務境界との対応表」です。

目的は `TaskPaneRefreshOrchestrationService` を分割することではありません。Phase 1 で inventory 化した R01-R16 について、理想的にはどの責務境界に所属するべきか、どこから安全に ownership 分離できるか、逆に何を orchestration に残すべきかを可視化します。

参照した正本:

- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/taskpane-display-recovery-current-state.md`
- `docs/taskpane-refresh-orchestration-responsibility-inventory.md`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshOrchestrationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WorkbookTaskPaneReadyShowAttemptWorker.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshPreconditionPolicy.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WindowActivatePaneHandlingService.cs`

今回も docs-only です。コード移動、service 新設、class rename、namespace 移動、retry 順序変更、trace 名変更、route 契約変更、fail-closed 条件変更、COM restore 順序変更、UI policy 変更、実装 refactor は行いません。

## 読み方

この文書でいう「理想 owner」は、次の変更理由境界を指します。現時点でその名前の service を作るという意味ではありません。

| 理想責務境界 | 意味 |
| --- | --- |
| display route / trigger observation boundary | raw reason、`TaskPaneDisplayRequest`、WindowActivate downstream trace を収束させる境界。 |
| refresh precondition / fail-closed policy boundary | refresh へ進めてよいかを副作用なしで判定する境界。 |
| refresh dispatch boundary | orchestration から refresh execution owner へ渡す境界。 |
| display protocol session boundary | created CASE display session、handoff、one-time completion を守る境界。 |
| ready-show attempt result boundary | 1 attempt の結果を protocol outcome へ渡す境界。 |
| retry / fallback ownership boundary | ready-show retry、pending retry、active CASE fallback の順序と状態を守る境界。 |
| workbook pane window resolve boundary | workbook から pane 対象 window を解決する境界。 |
| normalized outcome boundary | visibility / refresh source / rebuild fallback / foreground を completion 判定可能な outcome に変換する境界。 |
| foreground guarantee boundary | foreground obligation の decision / outcome と execution primitive を分離して扱う境界。 |
| timer lifecycle boundary | WinForms Timer の生成、停止、破棄を retry semantics と結び付けて扱う境界。 |

分離危険度は、Phase 4 で owner だけを切り出す場合の挙動差分リスクです。低は「まず docs/tests で固定すれば触りやすい」、中は「単独分離は可能だが周辺 contract 固定が必要」、高は「現時点では orchestration に残すべき」です。

## Rxx 対応表

| ID | 現在の owner | 理想 owner | lifecycle 上の位置 | orchestration 必須か | 単独分離可能か | 分離危険度 | 変更頻度 | 将来 service 化候補 | coordinator のまま残すべきか | policy object 向き | decision object 向き |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| R01 | `TaskPaneRefreshOrchestrationService.TryRefreshTaskPane(...)` | display route / trigger observation boundary | refresh path 最上流。WorkbookActivate / WindowActivate / Startup / retry から入る。 | 部分的に必須。route 収束と completion 接続は必要。 | 部分的。trace formatting / request normalization は候補。 | 中 | 中 | 候補。ただし route coordinator 寄り。 | はい。raw result を completion chain へ渡すため。 | いいえ | 部分的 |
| R02 | `TaskPaneRefreshPreconditionPolicy.DecideRefreshPrecondition(...)` | refresh precondition / fail-closed policy boundary | refresh entry 直後、dispatch 前。 | いいえ。呼び出しと skip outcome 接続は orchestration だが判定自体は外へ出ている。 | 完了済み。Phase 4 最初の safe unit として分離。 | 低から中 | 低 | 既存 policy を正本化済み。 | いいえ | はい | はい |
| R03 | `RefreshDispatchShell` | refresh dispatch boundary | precondition pass 後、normalized outcome 前。 | 部分的に必須。execution owner との bridge は必要。 | 部分的。API contract 固定後。 | 中 | 中 | 候補。薄い dispatch boundary。 | 部分的。completion 前の raw result 受け口は残る。 | いいえ | いいえ |
| R04 | `ShowWorkbookTaskPaneWhenReady(...)` | display protocol session boundary | `KernelCasePresentationService` の post-release 後、ready-show enqueue。 | はい。session start と completion owner が同一 protocol。 | いいえ。R14 と一体。 | 高 | 低から中 | 単独 service 化は後回し。 | はい | いいえ | いいえ |
| R05 | `WorkbookTaskPaneReadyShowAttemptWorker` + `HandleWorkbookTaskPaneShown(...)` | ready-show attempt result boundary | ready-show attempt が shown と判定された直後。 | はい。callback は outcome normalization と completion に接続する。 | いいえ。R10/R13/R14 と一体。 | 高 | 中 | attempt result adapter は候補だが後回し。 | はい | いいえ | はい |
| R06 | `TaskPaneReadyShowRetryScheduler` | retry / fallback ownership boundary | ready-show attempt 1 失敗後、attempt 2 を 80ms で schedule。 | いいえ。ただし順序 contract は orchestration が見る。 | 完了済み。R06 safe unit として scheduler ownership を分離。 | 低 | 低 | 既存 scheduler boundary。 | いいえ | いいえ | いいえ |
| R07 | `ScheduleWorkbookTaskPaneRefresh(...)` | retry / fallback ownership boundary | ready-show attempts exhausted 後の handoff と、workbook-target delayed refresh entry の二重性を持つ。 | はい。fallback handoff、immediate refresh、WorkbookOpen skip、pending retry entry、completion chain が近接する。 | いいえ。runtime extraction STOP。 | 高 | 中 | Phase 5 の protocol-preserving orchestration shrink 候補。 | はい | いいえ | はい |
| R08 | `PendingPaneRefreshRetryService` | retry / fallback ownership boundary | pending retry timer tick、workbook target / active CASE fallback。 | いいえ。既に分離境界がある。 | 部分的。file boundary は分離済みだが active fallback semantics の追加 runtime extraction は STOP。 | 中から高 | 中 | Phase 5 の retry convergence / display session / completion owner と一緒に扱う候補。 | いいえ | いいえ | 部分的 |
| R09 | `WorkbookPaneWindowResolver` | workbook pane window resolve boundary | ready-show attempt、fallback prepare、pending retry、coordinator ensure-window。 | いいえ。ただし activation policy は orchestration contract。 | 部分的。route 別 `activateWorkbook` 固定後。 | 中から高 | 中 | 候補。ただし UI helper ではない。 | 部分的 | いいえ | はい |
| R10 | `CompleteVisibilityRecoveryOutcome(...)` | normalized outcome boundary | skip / refresh / ready-show callback 後。 | 部分的に必須。completion 判定に使う。 | はい。ただし display-completable 固定後。 | 中 | 中 | 候補。outcome builder / decision object。 | 部分的 | いいえ | はい |
| R11 | `CompleteRefreshSourceSelectionOutcome(...)` | normalized outcome boundary | visibility outcome 後、rebuild fallback 前。 | いいえ。ただし completion trace と接続。 | はい。 | 中 | 中 | 候補。outcome builder。 | いいえ | いいえ | はい |
| R12 | `CompleteRebuildFallbackOutcome(...)` | normalized outcome boundary | refresh source outcome 後、foreground outcome 前。 | いいえ。ただし completion trace と接続。 | はい。 | 中 | 中 | 候補。outcome builder。 | いいえ | いいえ | はい |
| R13 | `CompleteForegroundGuaranteeOutcome(...)` + `TaskPaneRefreshCoordinator.ExecuteFinalForegroundGuaranteeRecovery(...)` | foreground guarantee boundary | visibility / source / rebuild outcome 後、completion 前。 | はい。decision / terminal outcome は必須。execution primitive は別 owner。 | いいえ。R14 / completion chain と強結合。foreground runtime extraction STOP。 | 高 | 中 | Phase 5 の foreground linkage / completion 接続と一緒に扱う候補。単独 service 化は後回し。 | はい | いいえ | はい |
| R14 | `BeginCreatedCaseDisplaySession(...)` / `TryCompleteCreatedCaseDisplaySession(...)` | display protocol session boundary | ready-show acceptance で start、ready-show callback or refresh path 終端で completion。 | はい。最重要 coordinator 残存領域。 | いいえ。 | 高 | 低から中 | 現時点では service 化しない。 | はい | いいえ | はい |
| R15 | `WindowActivatePaneHandlingService` + `WindowActivateDownstreamObservation` | display route / trigger observation boundary | WindowActivate dispatch 後、refresh entry の start / outcome。 | 部分的に必須。誤認防止 trace は近接が必要。 | 完了済み。R15 safe unit として downstream observation owner を分離。 | 中 | 中 | 既存 observation boundary。 | 部分的 | いいえ | はい |
| R16 | `TaskPaneRetryTimerLifecycle`。停止入口は `StopPendingPaneRefreshTimer(...)` | timer lifecycle boundary | success / shown callback / explicit stop。 | いいえ。 | はい。Phase 4 R16 safe unit で timer lifecycle owner を分離済み。 | 低 | 低 | 完了。timer lifecycle owner。 | いいえ | いいえ | いいえ |

## Phase 5 convergence boundary map

Phase 5 では、Rxx を単なる extraction candidate ではなく display protocol convergence topology として読み直します。

convergence topology:

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

| topology layer | 関連 R | current owner / input | Phase 5 固定 |
| --- | --- | --- | --- |
| raw facts | R05 / R07 / R08 / R09 / R15 | ready-show worker outcome、pending retry result、active fallback result、window resolve result、WindowActivate downstream observation。 | raw facts は completion ではない。callback、retry success、WindowActivate dispatch、window resolve success を direct completion と読まない。 |
| normalization | R10 / R11 / R12 | visibility / refresh source / rebuild fallback outcome mapping。 | normalized outcome は completion input になれるが completion owner ではない。 |
| foreground interpretation | R13 | foreground decision / execution raw result / `ForegroundGuaranteeOutcome`。 | foreground outcome は completion input だが emit owner ではない。`RequiredDegraded` は display-completable terminal であり success / failure へ丸めない。 |
| completion gate | R14 | created CASE display session、hard gate、one-time emit state。 | R14 だけが completion gate / one-time emit owner。 |
| one-time emit | R14 | `case-display-completed`。 | session ごとに 1 回だけ emit。worker / retry / WindowActivate / foreground / fallback へ ownership を戻さない。 |

### R05 callback / completion convergence freeze

R05 の callback は `shown raw facts` を orchestration convergence chain に戻す境界です。completion callback ではありません。

固定する読み方:

- `TaskPaneDisplayRetryCoordinator` は attempt sequencing owner です。
- `WorkbookTaskPaneReadyShowAttemptWorker` は ready-show attempt raw fact owner です。
- `HandleWorkbookTaskPaneShown(...)` は raw facts を R10/R11/R12/R13/R14 の chain に戻す convergence entry です。
- callback 後も R14 hard gate を満たすまで `case-display-completed` ではありません。

読み替え禁止:

- callback = display completed。
- callback = recovery completed。
- callback = foreground completed。
- callback = final success。

### R10/R11/R12 normalized outcome freeze

R10/R11/R12 は completion 判定可能な facts を作る normalized outcome boundary です。

固定する読み方:

- visibility outcome は pane visible / terminal / display-completable の input です。
- refresh source outcome は source / fallback / rebuild required の input です。
- rebuild fallback outcome は continuation 可否の input です。
- normalized outcome は completion ではありません。
- normalized outcome owner を one-time emit owner とみなしません。

### R13 foreground interpretation freeze

R13 は foreground outcome chain です。completion input ですが completion owner ではありません。

固定する読み方:

- foreground decision は refresh success、pane visible、refresh completed、foreground window、recovery service availability を見て required / not-required / degraded を解釈します。
- `final-foreground-guarantee-completed` は foreground execution completion observation であり、display completion ではありません。
- `RequiredSucceeded` は direct completion ではありません。
- `RequiredDegraded` は display-completable terminal ですが、success / failure へ丸めず、direct completion と読みません。

### R14 completion hard gate freeze

R14 は display session convergence の hard gate です。

固定する読み方:

- created CASE display session は ready-show acceptance で開始します。
- completion は `case-display-completed` の one-time emit だけです。
- emit は created CASE display reason、refresh success、pane visible、visibility terminal / display-completable、foreground terminal / display-completable、session が解決できること、session 未完了を満たす場合だけ成立します。
- R14 は state bag helper ではなく completion ownership boundary です。

### Non-completion owner map

| owner / boundary | completion ではないもの |
| --- | --- |
| `WorkbookTaskPaneReadyShowAttemptWorker` | `taskpane-already-visible`、`ready-show-attempt-result refreshed=true`、attempt shown raw facts。 |
| `PendingPaneRefreshRetryService` | `defer-retry-end refreshed=true`、`defer-active-context-fallback-end refreshed=true`、timer stop。 |
| `WorkbookPaneWindowResolver` / R09 | `resolve-window-success` 相当、`activateWorkbook=true` の activation request。 |
| `WindowActivatePaneHandlingService` / `WindowActivateDownstreamObservation` | `display-refresh-trigger-dispatched`、`window-activate-display-refresh-trigger-outcome`。 |
| R10/R11/R12 outcome mapping | normalized outcome。 |
| R13 foreground chain | `foreground-recovery-decision`、`final-foreground-guarantee-completed`、`RequiredSucceeded`、`RequiredDegraded`。 |

Phase 5 で redesign する場合も、この map を保った protocol-preserving convergence redesign に限定します。

### Phase 5 second runtime state boundary

Phase 5 第二実装後、runtime で完了済みの安全単位は次の 2 点です。

| completed safe unit | 意味 | owner 移動 |
| --- | --- | --- |
| R14 hard gate decision helper | visibility / foreground display-completable facts に基づく yes/no decision を private helper 化。 | なし。R14 completion owner は `TaskPaneRefreshOrchestrationService` に残る。 |
| R14 payload assembly helper | `case-display-completed` details payload assembly を private helper 化。 | なし。emit owner / trace owner / session owner は移動しない。 |

R14 で未移動のもの:

- completion owner。
- `case-display-completed` emit owner。
- one-time emit guard。
- display session lookup。
- `IsCompleted` guard。
- lock。
- `_createdCaseDisplaySessions` からの dictionary remove。
- `NewCaseVisibilityObservation.Complete(...)`。
- trace emit position。

この到達点は R14 の protocol を state bag helper へ移すものではありません。`BuildCaseDisplayCompletedDetailsPayload(...)` は string / details assembly だけを担い、`EvaluateCreatedCaseDisplayCompletionDecision(...)` は yes/no decision だけを担います。どちらも completion emit、session lifecycle、callback / pending / foreground / normalized outcome semantics を持ちません。

次 runtime 候補は R10/R11/R12 normalized outcome chain 呼び出し整理です。ただしこれはまだ GO ではなく、tests-first / safety net 評価を先に置きます。

現時点 STOP:

- display session lookup / one-time emit guard helper 化。
- callback raw facts adapter。
- foreground display-completable helper 化。
- route / dispatch shell 整理。
- R07/R09/R13/R14 をまたぐ runtime extraction。

### Phase 5 third runtime state boundary

Phase 5 第三実装後、runtime で完了済みの安全単位は次の 3 点です。

| completed safe unit | 意味 | owner 移動 |
| --- | --- | --- |
| R14 hard gate decision helper | visibility / foreground display-completable facts に基づく yes/no decision を private helper 化。 | なし。R14 completion owner は `TaskPaneRefreshOrchestrationService` に残る。 |
| R14 payload assembly helper | `case-display-completed` details payload assembly を private helper 化。 | なし。emit owner / trace owner / session owner は移動しない。 |
| R10/R11/R12 normalized outcome chain helper | `CompleteVisibilityRecoveryOutcome(...)` -> `CompleteRefreshSourceSelectionOutcome(...)` -> `CompleteRebuildFallbackOutcome(...)` の既存順序を private helper 化。 | なし。normalized outcome は completion owner ではない。 |

`CompleteNormalizedOutcomeChain(...)` は R10 -> R11 -> R12 の呼び出しだけを担います。completion 判定、foreground 判定、session lookup、one-time emit guard、`case-display-completed` emit、WindowActivate semantics、callback / pending の意味付けは持ちません。

第三実装後も未移動のもの:

- R13 foreground interpretation。
- R14 completion gate。
- `case-display-completed` emit owner。
- display session boundary。
- session lookup。
- `IsCompleted` guard。
- lock。
- `_createdCaseDisplaySessions` からの dictionary remove。
- `NewCaseVisibilityObservation.Complete(...)`。
- WindowActivate handling。
- trace owner / payload contract。

trace 名、trace source、trace payload field set / order / names / values は維持します。normal refresh / ready-show callback / precondition skip の各 path は、R10/R11/R12 normalized outcome chain と R13/R14 との距離を変えません。

次 runtime 候補はまだ GO ではありません。次に runtime を触る場合は、改めて tests-first / safety net 評価を先に置きます。

現時点 STOP:

- foreground display-completable 判定 helper 化。
- display session lookup / one-time emit guard helper 化。
- callback raw facts adapter。
- route / dispatch shell 整理。
- R07/R09/R13/R14 横断 extraction。

### Phase 5 R13 trace details runtime state boundary

Phase 5 R13 foreground trace details helper 化後、runtime で完了済みの安全単位は次の 5 点です。

| completed safe unit | 意味 | owner 移動 |
| --- | --- | --- |
| R14 hard gate decision helper | visibility / foreground display-completable facts に基づく yes/no decision を private helper 化。 | なし。R14 completion owner は `TaskPaneRefreshOrchestrationService` に残る。 |
| R14 payload assembly helper | `case-display-completed` details payload assembly を private helper 化。 | なし。emit owner / trace owner / session owner は移動しない。 |
| R10/R11/R12 normalized outcome chain helper | `CompleteVisibilityRecoveryOutcome(...)` -> `CompleteRefreshSourceSelectionOutcome(...)` -> `CompleteRebuildFallbackOutcome(...)` の既存順序を private helper 化。 | なし。normalized outcome は completion owner ではない。 |
| R13 classification helper | foreground execution result を `RequiredSucceeded` / `RequiredDegraded` に分類する処理を private helper 化。 | なし。R13 内の局所 helper 化であり、foreground execution / trace / completion owner は移動しない。 |
| R13 trace details helper | foreground observation details payload assembly を private helper 化。 | なし。details 文字列 assembly のみであり、trace emit / foreground execution / completion owner は移動しない。 |

`ClassifyRequiredForegroundExecutionOutcome(...)` は execution result から foreground outcome を返す分類だけを担います。`ExecutionAttempted && Recovered` の場合だけ `RequiredSucceeded` を返し、それ以外は現行通り `RequiredDegraded` を返します。`RequiredDegraded` を `RequiredFailed`、success、failure、direct completion へ丸めません。

R13 trace details helper の意味:

- `BuildForegroundRecoveryDecisionDetails(...)` は `foreground-recovery-decision` details の文字列 assembly だけを担う。field set / order は `reason -> foregroundRecoveryStarted -> foregroundSkipReason -> foregroundOutcomeStatus`。
- `BuildFinalForegroundGuaranteeStartedDetails(...)` は `final-foreground-guarantee-started` details の文字列 assembly だけを担う。field set / order は `reason`。
- `BuildFinalForegroundGuaranteeCompletedDetails(...)` は `final-foreground-guarantee-completed` details の文字列 assembly だけを担う。field set / order は `reason -> recovered -> foregroundOutcomeStatus`。
- completed mapping は `recovered=true` の場合だけ `RequiredSucceeded`、`recovered=false` の場合は `RequiredDegraded`。
- `RequiredDegraded` は `RequiredFailed`、success、direct completion へ丸めない。

R13 classification helper が持たないもの:

- foreground execution 呼び出し。
- foreground trace emit。
- WindowActivate handling。
- R14 completion gate。
- `case-display-completed` emit。
- session lookup。
- one-time emit guard。
- callback / pending / normalized outcome の意味付け。

R13 trace details helper 化後も未移動のもの:

- foreground execution 呼び出し。
- trace action / source / emit position。
- logger action / 発火順。
- WindowActivate handling。
- R14 completion gate。
- `case-display-completed` emit owner。
- display session boundary。
- session lookup。
- `IsCompleted` guard。
- lock。
- `_createdCaseDisplaySessions` からの dictionary remove。
- `NewCaseVisibilityObservation.Complete(...)`。
- trace owner / payload contract。

trace 名、trace source、trace payload field set / order / names / values は維持します。foreground outcome != completion、`RequiredDegraded` は success / failure / direct completion ではない、`RequiredSucceeded` は input only、`RequiredFailed` は completion gate を通さない、`NotRequired` / `SkippedAlreadyVisible` は foreground success ではない、callback != completion、pending != completion、WindowActivate dispatch != completion、`case-display-completed` one-time emit、display session boundary、foreground outcome semantics は維持します。

次 runtime 候補はまだ GO ではありません。次に runtime を触る場合は、foreground display-completable 判定 helper 化の tests-first 評価を先に置きます。

現時点 STOP:

- foreground display-completable helper 化。
- display session lookup / one-time emit guard helper 化。
- callback raw facts adapter。
- route / dispatch shell 整理。
- R07/R09/R13/R14 横断 extraction。

### Phase 5 R13 foreground display-completable runtime state boundary

Phase 5 foreground display-completable helper 化後、runtime で完了済みの安全単位は次の 6 点です。

| completed safe unit | 意味 | owner 移動 |
| --- | --- | --- |
| R14 hard gate decision helper | visibility / foreground display-completable facts に基づく yes/no decision を private helper 化。 | なし。R14 completion owner は `TaskPaneRefreshOrchestrationService` に残る。 |
| R14 payload assembly helper | `case-display-completed` details payload assembly を private helper 化。 | なし。emit owner / trace owner / session owner は移動しない。 |
| R10/R11/R12 normalized outcome chain helper | `CompleteVisibilityRecoveryOutcome(...)` -> `CompleteRefreshSourceSelectionOutcome(...)` -> `CompleteRebuildFallbackOutcome(...)` の既存順序を private helper 化。 | なし。normalized outcome は completion owner ではない。 |
| R13 classification helper | foreground execution result を `RequiredSucceeded` / `RequiredDegraded` に分類する処理を private helper 化。 | なし。foreground execution / trace / completion owner は移動しない。 |
| R13 trace details helper | foreground observation details payload assembly を private helper 化。 | なし。details 文字列 assembly のみであり、trace emit / foreground execution / completion owner は移動しない。 |
| R13 foreground display-completable input helper | R14 hard gate が読む foreground input 判定を private helper 化。 | なし。completion owner / emit owner / session owner は移動しない。 |

`IsForegroundDisplayCompletableTerminalInput(...)` は `outcome != null`、`outcome.IsTerminal`、`outcome.IsDisplayCompletable` だけを読みます。これは display-completable input 判定であり、completion 判定全体ではありません。

display-completable terminal input:

- `RequiredSucceeded`
- `RequiredDegraded`
- `NotRequired`
- `SkippedAlreadyVisible`

non-display-completable input:

- `RequiredFailed`
- `SkippedNoKnownTarget`
- `Unknown`

この mapping は `display-completable input != completion` を前提にします。`RequiredDegraded` は success / failure / direct completion へ丸めません。`NotRequired` / `SkippedAlreadyVisible` は foreground success ではありません。

foreground display-completable helper が持たないもの:

- completion 判定全体。
- `case-display-completed` emit。
- session lookup。
- one-time emit guard。
- `IsCompleted` guard。
- lock。
- `_createdCaseDisplaySessions` からの dictionary remove。
- foreground execution。
- WindowActivate handling。
- trace emit。
- callback / pending の意味付け。

foreground display-completable helper 化後も未移動のもの:

- R14 completion gate。
- `case-display-completed` emit owner。
- display session boundary。
- session lookup。
- `IsCompleted` guard。
- lock。
- `_createdCaseDisplaySessions` からの dictionary remove。
- `NewCaseVisibilityObservation.Complete(...)`。
- foreground execution。
- WindowActivate handling。
- trace owner / payload contract。

foreground outcome != completion、display-completable input != completion、`RequiredDegraded` は success / failure / direct completion ではない、`RequiredSucceeded` は input only、`RequiredFailed` は completion gate を通さない、`NotRequired` / `SkippedAlreadyVisible` は foreground success ではない、callback != completion、pending != completion、WindowActivate dispatch != completion、`case-display-completed` one-time emit、display session boundary、trace contract、foreground outcome semantics は維持します。

現時点 STOP:

- display session lookup / one-time emit guard helper 化。
- callback raw facts adapter。
- route / dispatch shell 整理。
- R07/R09/R13/R14 横断 extraction。

次はすぐ runtime へ進まず、Phase 5 runtime の一区切り判断を docs-only / read-only で行う候補があります。

### Phase 5 runtime closure boundary

Phase 5 runtime は、R13 foreground display-completable input helper 化後の時点で一度区切ります。

ここまでで完了した safe unit は、すべて private / pure / local helper として切れる範囲に限定されています。

| completed safe unit | closure 上の意味 | owner 移動 |
| --- | --- | --- |
| R14 completion hard gate decision helper | completion hard gate の yes/no decision を局所化。 | なし。R14 completion owner は残る。 |
| R14 `case-display-completed` payload helper | completion trace details assembly を局所化。 | なし。emit owner / trace owner / session owner は残る。 |
| R10/R11/R12 normalized outcome chain helper | R10 -> R11 -> R12 の呼び出し順を局所化。 | なし。normalized outcome は completion owner ではない。 |
| R13 foreground classification helper | execution result から `RequiredSucceeded` / `RequiredDegraded` への分類を局所化。 | なし。foreground execution / trace / completion owner は残る。 |
| R13 foreground trace details helpers | foreground observation details assembly を局所化。 | なし。trace emit owner / foreground execution owner は残る。 |
| R13 foreground display-completable input helper | R14 hard gate が読む foreground input 判定を局所化。 | なし。completion owner / emit owner / session owner は残る。 |

完了済み helper 群は、completion owner、emit owner、session lifecycle、callback meaning、pending retry semantics、WindowActivate semantics、trace contract を持ちません。

未移動 owner:

- R14 completion gate。
- `case-display-completed` emit owner。
- display session boundary。
- session lookup。
- `IsCompleted` guard。
- lock。
- `_createdCaseDisplaySessions` からの dictionary remove。
- `NewCaseVisibilityObservation.Complete(...)`。
- callback meaning。
- pending retry。
- WindowActivate handling。
- trace owner / payload contract。

STOP 継続:

- display session lookup / one-time emit guard helper 化。
- callback raw facts adapter。
- route / dispatch shell 整理。
- R07/R09/R13/R14 横断 extraction。

これらは completion owner、session lifecycle、callback meaning、route semantics、trace contract に近い領域です。現時点で runtime extraction すると、protocol-preserving helper 化ではなく protocol rewrite になりやすいため STOP を継続します。

次に進む場合は、runtime 実装から始めません。read-only 棚卸し、tests-first 評価、docs freeze を先に行い、その後に freeze line を変えない最小単位だけを改めて GO / STOP 判定します。

closure 後も immutable freeze line として維持するもの:

- foreground outcome != completion。
- display-completable input != completion。
- normalized outcome != completion。
- callback != completion。
- pending != completion。
- WindowActivate dispatch != completion。
- `case-display-completed` one-time emit。
- display session boundary。
- trace contract。
- trace payload field set / order / names / values。
- retry sequencing。
- foreground outcome semantics。
- `RequiredDegraded` は success / failure / direct completion ではない。

## R07 runtime extraction STOP

R07 は、現時点では runtime extraction を行いません。`ScheduleWorkbookTaskPaneRefresh(...)` は単なる delayed timer schedule helper ではなく、ready-show fallback handoff trace、`WorkbookOpen` skip、workbook target tracking、window resolve、pre-timer immediate refresh、pending retry start decision を 1 つの protocol entry として束ねています。

現行 caller / reason には、created CASE ready-show exhaustion 後の handoff と、`KernelHomeForm.OpenSheet.PostClose` の workbook-target delayed refresh entry が含まれます。この二重性があるため、R07 を ready-show exhaustion 専用 owner として切り出すと protocol rewrite に化けやすいです。

immediate refresh success と pending retry success は completion ではありません。どちらも existing refresh / outcome / completion chain への re-entry であり、`case-display-completed` は display session owner の条件を満たした場合だけ emit できます。

`WorkbookOpen` skip は null guard ではなく window stability boundary の runtime stabilization contract です。handoff trace 後でも `WorkbookOpen` 直後の unresolved window では pending retry start へ進まず、後続の `WorkbookActivate` / `WindowActivate` 側へ委ねます。

`PendingPaneRefreshRetryService` の active CASE context fallback は、tracked workbook を見失った場合の target-lost resiliency fallback です。completion fallback ではなく、成功しても completion owner は orchestration 側に残ります。

R07 は Phase 5 で、protocol-preserving orchestration shrink として扱う候補です。Phase 4 では R07 runtime separation は STOP とし、R06/R08/R14/R10-R13 との freeze line を壊さないことを優先します。

## R08 active fallback semantics STOP / truth table

`PendingPaneRefreshRetryService` は既に file boundary として分離されています。ただし active CASE fallback semantics は、現時点では追加 runtime extraction や semantics 変更を行いません。

R08 active fallback が持つ current-state の意味:

- pending retry `400ms / 3 attempts` の tick で、まず tracked workbook を探す。
- tracked workbook が見つかった場合は workbook target refresh を試す。
- tracked workbook を見失った場合だけ active context を解決し、active context が CASE なら `TryRefreshTaskPane(reason, null, null)` で active refresh fallback を試す。
- tracked workbook も active CASE context もない場合、または attempts exhausted の場合は timer を止める。
- tracked workbook route / active CASE fallback route の refresh success 時も timer を止める。

R08 active fallback が持たない意味:

- completion fallback ではありません。
- recovered event ではありません。
- foreground fallback / foreground success ではありません。
- display session completion ではありません。
- `case-display-completed` emit owner ではありません。

truth table:

| tracked workbook exists | active context is CASE | attempts remaining | refresh attempted | refresh target | timer continues | completion meaning | trace / outcome meaning |
| --- | --- | --- | --- | --- | --- | --- | --- |
| true | 該当なし | yes | yes | tracked workbook + resolved pane window | refresh success なら stop。refresh failure かつ attempts が残る場合だけ継続。 | active fallback 自体は completion を emit しない。tracked refresh success も completion ではない。 | `defer-retry-start` / `defer-retry-end` は workbook-target retry attempt の観測。 |
| false | true | yes | yes | active CASE context | refresh success なら stop。refresh failure かつ attempts が残る場合だけ継続。 | active fallback 自体は completion を emit しない。active CASE fallback success も completion ではない。 | `defer-active-context-fallback-start` / `defer-active-context-fallback-end` は target-lost resiliency fallback の観測。 |
| false | false | yes | no | none | stop | active fallback 自体は completion を emit しない。stop は completion ではない。 | `defer-active-context-fallback-stop` は fallback 不成立の観測。 |
| any | any | no | no | none | stop | active fallback 自体は completion を emit しない。attempts exhausted は completion ではない。 | attempts exhausted による stop は retry lifecycle の観測。 |

Phase 5 boundary:

- active fallback semantics は Phase 5 で retry convergence、display session、completion owner、foreground linkage と一緒に扱う候補です。
- Phase 4 では active fallback runtime extraction、service 新設、trace 名 / trace payload / trace source 変更、retry semantics 変更、completion semantics 変更、foreground semantics 変更を行いません。
- active CASE fallback を completion helper、foreground helper、recovered-event helper、display-session helper として切り出しません。
- R08 を触る場合も、R07 handoff、R14 display session、R13 foreground outcome、R10/R11/R12 normalized outcome との接続を同時に確認するまで semantics を変えません。

## R09 runtime extraction STOP / route matrix

R09 は `workbook pane window resolve boundary` ですが、現時点では runtime extraction を行いません。`ResolveWorkbookPaneWindow(...)` は単なる window getter ではなく、route ごとの `activateWorkbook` side effect と window availability boundary を含むためです。

R09 が持つ current-state の意味:

- 対象 workbook の first visible window、または active workbook が対象 workbook と一致する場合の active window を pane 対象 window として解決する。
- `activateWorkbook=true` の route だけ `ExcelInteropService.ActivateWorkbook(workbook)` を呼ぶ。
- window unresolved を推測で補わず、caller route の retry / fallback / fail-closed protocol へ返す。

R09 が持たない意味:

- completion 判定。
- foreground decision。
- retry success。
- fallback start / stop。
- `case-display-completed` emit。

route 別の freeze line:

| route | `activateWorkbook` | boundary owner として固定する意味 |
| --- | --- | --- |
| ready-show immediate / wait-ready path | `true` | ready-show attempt の window availability を確認する。未解決なら attempt failure 側へ戻り、`attempt 1 -> 80ms attempt 2 -> pending fallback` の順序に従う。 |
| R07 `ScheduleWorkbookTaskPaneRefresh(...)` immediate refresh | `false` | pending timer 開始前の immediate refresh input を作る。未解決でも completion / retry success ではなく、refresh path と pending decision に委ねる。 |
| R08 pending retry tick | `true` | tracked workbook に対する retry attempt のための window resolve / activation request。foreground guarantee success ではない。 |
| `WorkbookOpen` skip / stabilization boundary | 該当なし | `TaskPaneRefreshPreconditionPolicy` が R09 前で止める。skip は stabilization contract であり pending start / completion ではない。 |
| `WindowActivate` downstream observation | event window があれば R09 は不要。補完時は `false` | event window / dispatch / downstream trace は observation。`Dispatched != completion` を維持する。 |
| `KernelHomeForm.OpenSheet.PostClose` delayed refresh | immediate prepare は `false`、pending tick は `true` | workbook-target delayed refresh entry として R07/R08 に従う。created CASE completion route ではない。 |
| foreground guarantee path | 該当なし | R13 の decision / outcome chain。R09 の activation request や resolved window と foreground terminal を同一視しない。 |

R09 runtime extraction 禁止:

- Phase 4 では R09 の service extraction、class rename、namespace 移動、route policy 移動を行いません。
- Phase 5 で扱う場合も、`activateWorkbook` route matrix、`WorkbookOpen` stabilization boundary、pending retry / WindowActivate / foreground / completion との距離を壊さない protocol-preserving shrink に限定します。
- R09 を UI helper、foreground helper、retry helper、completion helper として切り出しません。

## R13 foreground outcome chain runtime extraction STOP / decision contract

R13 は foreground guarantee boundary ですが、現時点では runtime extraction を行いません。foreground outcome chain は foreground decision / outcome / trace と execution bridge を扱う decision contract であり、display session completion owner そのものではありません。

R13 が持つ current-state の意味:

- refresh success と pane visible が揃った後にだけ foreground decision を評価する。
- refresh completed、foreground window、foreground recovery service が揃う場合だけ foreground execution を required path として試す。
- execution bridge は `TaskPaneRefreshCoordinator.ExecuteFinalForegroundGuaranteeRecovery(...)` で、raw recovered fact を返す。
- decision step は raw execution result を `RequiredSucceeded`、`RequiredDegraded`、`NotRequired` などの `ForegroundGuaranteeOutcome` に正規化する。
- terminal / display-completable な foreground outcome は completion input になれる。

R13 が持たない意味:

- `case-display-completed` emit owner ではありません。
- display session boundary ではありません。
- visibility recovery outcome owner ではありません。
- WindowActivate dispatch / downstream observation owner ではありません。
- retry success / pending fallback owner ではありません。
- foreground execution success を display completion に直結させる owner ではありません。

主要 outcome の fixed meaning:

| outcome | fixed meaning | 読み替え禁止 |
| --- | --- | --- |
| `RequiredSucceeded` | foreground guarantee が必要で、execution が attempted かつ recovered。display-completable terminal。 | 単独で `case-display-completed` ではない。 |
| `RequiredDegraded` | foreground guarantee が必要だったが、recovered false。display-completable terminal だが success guarantee ではない。 | success / failure のどちらにも雑に丸めない。 |
| `NotRequired` | foreground guarantee が不要、または required path の事実が揃わない。display-completable terminal。 | foreground success ではなく foreground path 非該当。 |

Phase 5 boundary:

- foreground outcome chain は R13 / Phase 5 で foreground linkage、retry convergence、display session、completion owner と一緒に扱う候補です。
- Phase 4 では foreground runtime extraction、service 新設、trace 名 / trace payload / trace source 変更、foreground semantics 変更、completion semantics 変更を行いません。
- visibility outcome と foreground outcome を同じ enum / 同じ success-failure meaning に統合しません。
- `RequiredDegraded` を success / failure に丸める decision object、helper、trace formatter を作りません。
- foreground outcome chain を UI helper、WindowActivate helper、retry helper、completion helper として切り出しません。

## Coupling Matrix

| ID | trace coupling | retry coupling | window dependency coupling | fail-closed coupling | display-session coupling | WindowActivate coupling |
| --- | --- | --- | --- | --- | --- | --- |
| R01 | 高。start/end と WindowActivate outcome を持つ。 | 低。ただし retry から再入される。 | 中。input window を route facts として保持。 | 中。skipped result を success にしない。 | 中。created CASE reason の completion check へ進む。 | 高。R15 と対になる。 |
| R02 | 中。skip action 名が trace / outcome source になる。 | 低。retry 中でも同じ gate。 | 中。WorkbookOpen window null が条件。 | 高。最初の fail-closed boundary。 | 低 | 中。protection gate と隣接。 |
| R03 | 中。raw result が end trace と outcome へ入る。 | 低 | 中。coordinator 側 window resolve に接続。 | 中。missing dependency / suppression / context reject を success にしない。 | 中。raw result が completion 入力。 | 低 |
| R04 | 高。enqueue / session start / handoff trace。 | 中。worker retry へ渡す。 | 中。入口 window は未確定。 | 中。session 開始条件が fail-closed。 | 高。同一 session の開始点。 | 低 |
| R05 | 高。attempt result と completion trace。 | 中。失敗時は retry / fallback 側へ分岐。 | 高。attempt resolved window を消費。 | 高。raw outcome 不足なら completion しない。 | 高。completion callback。 | 低 |
| R06 | 中。retry scheduled / firing。 | 高。ready-show attempt 2 の発火点。 | 低。retry action 内で再 resolve。 | 低。action null なら何もしない。 | 中。attempt result へ戻る。 | 低 |
| R07 | 高。fallback handoff / immediate / scheduled。 | 高。pending retry への入口。 | 高。handoff 前に window resolve。 | 高。WorkbookOpen skip なら fallback 開始しない。 | 高。ready-show failure を session completion へ戻す迂回路。 | 低 |
| R08 | 高。defer retry / active fallback trace。 | 高。400ms / 3 attempts の本体。 | 高。target retry では activateWorkbook=true。 | 中。target unresolved + active context not CASE で stop。 | 中。成功時は display completion chain へ戻る。 | 低 |
| R09 | 高。resolve-window-* trace。 | 中。retry / ready-show / coordinator から呼ばれる。 | 高。中心責務。 | 高。未解決を推測で補わない。 | 中。resolved window は completion details に効く。 | 中。event window と activation primitive を混同しやすい。 |
| R10 | 高。visibility-recovery-* trace。 | 中。retry result を消費する。 | 中。pane visible / foreground window facts を読む。 | 高。insufficient facts を display-completable にしない。 | 高。completion 条件そのもの。 | 低 |
| R11 | 高。refresh-source-* trace。 | 低 | 中。snapshot result は context/window に依存。 | 中。failed / not reached を success にしない。 | 中。completion details に含まれる。 | 低 |
| R12 | 高。rebuild-fallback-* trace。 | 低 | 低から中。snapshot/cache 経由。 | 中。failed / cannot continue を success にしない。 | 中。completion details に含まれる。 | 低 |
| R13 | 高。foreground decision / final guarantee trace。 | 低 | 高。foreground window と recovery primitive が条件。 | 高。required 条件を満たさなければ NotRequired。 | 高。terminal / display-completable が completion 条件。 | 中。WindowActivate 発火を terminal と誤認しやすい。 |
| R14 | 高。session start / handoff / completed trace。 | 中。retry result を消費。 | 中。workbook full name / window descriptor を使う。 | 高。条件未充足なら emit しない。 | 高。中心境界。 | 低から中。WindowActivate refresh でも created-case reason ならここへ来る。 |
| R15 | 高。誤認防止 trace 自体が責務。 | 低 | 高。event window と downstream resolved facts を区別。 | 高。Dispatched を completion とみなさない。 | 低 | 高。中心責務。 |
| R16 | 低。周辺 trace で観測。 | 高。retry lifecycle を止める。 | 低 | 低。timer absent は no-op。 | 中。shown callback / success 時に session 周辺を止める。 | 低 |

## Helper にしてはいけない理由 / 今は動かしてはいけない理由

| ID | helper にしてはいけない理由 | 今は動かしてはいけない理由 |
| --- | --- | --- |
| R01 | route normalization は単なる文字列整形ではなく、structured request、raw reason、trace、completion input を束ねるため。 | WindowActivate trace と `try-refresh-end` の意味が崩れると、dispatch と display success の誤認が再発する。 |
| R02 | pure policy へ寄せられるが、protection 判定は active window に依存する bridge を含むため単純 helper ではない。 | protection gate の意図が Phase 3 で freeze される前に動かすと、止める範囲が変わる。 |
| R03 | dispatch shell は lower owner の raw result を normalized outcome chain へ接続する境界で、単なる pass-through ではない。 | coordinator API / suppression count / Kernel HOME visible condition の変更と混ざりやすい。 |
| R04 | ready-show acceptance は session start と handoff trace を作る protocol entry であり、enqueue helper ではない。 | R14 と離すと `display-handoff-completed` と `case-display-completed` の対応が失われる。 |
| R05 | callback は worker の raw result を completion 判定可能な outcome に変換する入口で、attempt helper ではない。 | raw refresh success を直接 completion と読む実装へ戻りやすい。 |
| R06 | timer scheduling は小さいが、attempt 上限と retry delay の protocol 値を持つため汎用 timer helper ではない。 | `attempt 1 -> 80ms attempt 2` の順序が freeze される前に抽出すると retry semantics が薄まる。 |
| R07 | fallback handoff は ready-show failure を pending retry protocol へ変換する境界で、単なる schedule helper ではない。 | immediate refresh、WorkbookOpen skip、pending retry 開始の順序を崩すと表示不安定に直結する。 |
| R08 | active CASE context fallback を持つ retry owner であり、汎用 retry helper ではない。 | target workbook lost 時の fallback 条件が未固定のまま外へ出ると、誤った CASE へ refresh しやすい。 |
| R09 | `activateWorkbook` が副作用を持つため、window getter helper ではない。 | route ごとの activation 可否を固定する前に動かすと WorkbookOpen / WindowActivate 境界を壊す。 |
| R10 | visibility outcome は display-completable を決める decision で、trace formatter ではない。 | degraded / terminal の意味が completion 条件に直結する。 |
| R11 | source outcome は snapshot / cache / MasterList rebuild の観測契約で、文字列化 helper ではない。 | refresh source failure を success に丸める危険がある。 |
| R12 | rebuild fallback outcome は continuation 可否を表す decision で、log helper ではない。 | rebuild fallback を completion 条件に含める読み方がまだ固定対象。 |
| R13 | foreground decision は UI primitive 実行可否と terminal outcome を決めるため、UI helper ではない。 | execution owner と decision owner を混ぜると `RequiredDegraded` の意味が割れる。 |
| R14 | display session は one-time completion emit の owner であり、state bag helper ではない。 | ここを動かすと completion owner が分散し、already-visible path と refresh path の収束が壊れる。 |
| R15 | WindowActivate downstream trace は observation contract であり、log helper ではない。 | `WindowActivateDispatchOutcome.Dispatched` を completion と誤認しない線がまだ重要。 |
| R16 | cleanup は retry lifecycle と session completion の重複発火防止であり、Dispose helper ではない。 | R06/R08 の retry semantics / freeze line が固定される前に分けると timer leak / duplicate retry の責任が曖昧になる。 |

## 安全に触れそうな領域

Phase 4 safe-first の候補は、単独の挙動変更を避けやすく、先に docs / tests で固定できる領域です。

1. R02 refresh precondition / fail-closed policy boundary
   `WorkbookOpen` window-dependent skip は既に `TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(...)` にあり、pure 判定に近いです。Phase 4 最初の safe unit として、判定 owner は `TaskPaneRefreshPreconditionPolicy.DecideRefreshPrecondition(...)` へ分離済みです。

2. R16 timer lifecycle boundary
   no-op 条件が明確で、cleanup 自体は挙動を持ちにくいです。R06/R08 の retry semantics と freeze line 固定後なら安全度が高いです。

   Phase 4 R16 safe unit では `TaskPaneRetryTimerLifecycle` を timer lifecycle boundary とし、ready-show retry timer と pending retry timer の create / register / stop / unregister / dispose を寄せました。`TaskPaneRefreshOrchestrationService` は ready-show retry の schedule 順序と callback 接続を保持し、`PendingPaneRefreshRetryService` は pending retry callback / attempts / active CASE fallback の owner のままです。`attempt 1 -> 80ms retry attempt 2 -> pending retry fallback`、pending retry `400ms / 3 attempts`、completion 条件、display session boundary、trace 名、WindowActivate downstream と completion trace の境界は変更していません。

3. R06 ready-show retry timer
   `80ms` delay と attempt 2 発火に閉じた小さい境界です。R06 safe unit では `TaskPaneReadyShowRetryScheduler` を追加し、scheduled / firing trace と timer schedule decision を orchestration から分離しました。`TaskPaneDisplayRetryCoordinator` は attempt sequencing、`WorkbookTaskPaneReadyShowAttemptWorker` は attempt 本体、`TaskPaneRetryTimerLifecycle` は timer create / stop / dispose の owner のままです。`attempt 1 -> 80ms retry attempt 2 -> pending retry fallback`、pending retry `400ms / 3 attempts`、fallback handoff、callback、completion、display session、trace 名と意味は変更していません。

4. R10/R11/R12 normalized outcome mapping
   raw result から outcome を作る decision object として扱えます。display-completable / terminal の意味を tests で固定することが前提です。

5. R15 WindowActivate downstream trace
   event capture owner とは別の observation boundary として整理できます。ただし `Dispatched != completion` の contract を先に強く固定します。

## 後回しにすべき危険領域

次は Phase 2 では動かさず、Phase 3 で変更禁止条件を固定してから扱います。

- R04 + R14
  ready-show acceptance と display session completion は同一 protocol の開始と終了です。ここを離すと `display-handoff-completed` と `case-display-completed` の相関が壊れます。

- R05 + R10/R13/R14
  ready-show callback は raw attempt result を normalized outcomes へ変換し、completion 判定へ渡します。worker 側や coordinator 側へ戻すと raw refresh success を completion と誤読しやすいです。

- R06/R07/R08 retry sequence
  `attempt 1 -> 80ms retry attempt 2 -> pending retry fallback` は単なる再試行ではなく、window unresolved / flicker を避ける安全装置です。順序を固定するまでは一体で扱います。

- R09 activation policy
  `activateWorkbook=true/false` は route ごとの副作用です。window resolver を外へ出す場合でも、WorkbookOpen、ready-show、pending retry、WindowActivate の呼び分けを先に凍結します。

- R13 foreground guarantee decision
  `ExcelWindowRecoveryService` は execution primitive ですが、required / not-required / degraded / display-completable の decision は completion 条件です。UI helper 化してはいけません。現時点では foreground runtime extraction も行いません。

- R14 completion emit
  `case-display-completed` の唯一 emit owner です。Phase 4 の最後に扱うべきです。

## Orchestration に残すべき境界

現時点で coordinator のまま残る可能性が高い領域は次です。

- display protocol session boundary: R04 / R14。
- ready-show callback から completion への収束: R05 / R10 / R13 / R14。
- foreground terminal outcome と completion 判定の接続: R13 / R14。
- route / trigger observation と completion 誤認防止: R01 / R15。
- retry fallback が display session へ戻る収束点: R07 / R08 / R14。

理由は、これらが「何かを実行する処理」ではなく、複数 owner の raw facts を `case-display-completed` へ収束させる protocol owner だからです。

## 強結合になっている理由

### retry / fallback / normalized outcome

retry / fallback / normalized outcome は、同じ display completion を別経路から満たしにいくため強結合です。

- ready-show attempt は already-visible path と refresh path の両方を成功相当にできます。
- ready-show が失敗した場合だけ pending retry fallback へ落ちます。
- pending retry は target workbook を追い、見失った場合は active CASE context fallback を使います。
- どの経路で戻ってきても、completion は raw result ではなく visibility / refresh source / rebuild fallback / foreground の normalized outcomes を見ます。
- そのため retry owner だけを切ると、失敗・再試行・fallback・completion の意味が分散します。

安全な読み方は、R06/R07/R08 を retry sequence として固定し、R10/R11/R12/R13 を normalized outcome boundary として固定し、その上で R14 が final completion を一元的に判断する形です。

### WindowActivate downstream trace

WindowActivate downstream trace は completion と誤認されやすいです。

- `WindowActivatePaneHandlingService` の `Dispatched` は display request を投げたことだけを示します。
- downstream の `window-activate-display-refresh-trigger-start` / `outcome` は、WindowActivate 由来の refresh path がどう終わったかを観測する trace です。
- どちらも `case-display-completed` の成立そのものではありません。
- `WindowActivate` は window-safe な trigger ですが、recovery owner、foreground owner、display completion owner ではありません。

そのため R15 は log helper ではなく observation boundary として扱います。R01 と近い位置に残るのは、trigger と refresh outcome を並べて観測し、`Dispatched == display success` という誤読を防ぐためです。

### display-session boundary

display-session boundary は簡単に切れません。

- session start は ready-show acceptance の時点で、workbook full name と created-case reason に紐づきます。
- completion は ready-show callback だけでなく、refresh path 終端からも来ます。
- already-visible path と refresh path の両方が同じ session に収束します。
- `case-display-completed` は one-time emit で、重複 emit を防ぐ状態管理を持ちます。
- foreground terminal / visibility terminal が揃うまで completion しません。

したがって R04/R14 は Phase 4 後半まで coordinator に残すべきです。state bag として helper 化すると、completion owner が見えなくなります。

### ready-show callback

ready-show callback が orchestration に残りやすい理由は、worker の責務が attempt 実行で止まるためです。

- worker は window resolve、already-visible 判定、refresh delegate 呼び出しを行います。
- callback で返る outcome はまだ display protocol completion ではありません。
- orchestration はその outcome に visibility / refresh source / rebuild fallback / foreground を補完し、display-session completion を判定します。

つまり callback は worker の終了点ではなく、display protocol の再収束点です。ここを lower worker 側へ戻すと、worker が completion owner になってしまいます。

### foreground decision

foreground decision は UI helper 化できません。

- `ExcelWindowRecoveryService` は workbook/window recovery と foreground promotion の execution primitive です。
- `TaskPaneRefreshCoordinator` は execution bridge を持ちます。
- しかし `RequiredSucceeded`、`RequiredDegraded`、`NotRequired`、`SkippedAlreadyVisible` を display-completable な terminal outcome として読む責務は display protocol 側です。
- foreground recovery を試みる条件は `refresh succeeded`、`pane visible`、`refresh completed`、`window present`、`recovery service available` の組み合わせで決まります。

この decision を UI helper にすると、前面化の実行可否と display completion 条件が混ざります。R13 は decision object 向きですが、completion 接続は orchestration に残すべきです。

## Phase 4 safe-first 候補

Phase 4 の候補順は次です。ここでは service 新設を前提にせず、まず docs / tests / contract 固定で扱います。

1. R02 の boundary tests と fail-closed 固定は Phase 4 最初の safe unit として完了済み。
2. R16 の timer lifecycle owner 固定は完了済み。
3. R06 の ready retry scheduler 固定は safe unit として完了済み。
4. R10/R11/R12 の normalized outcome mapping は完了済み。
5. R15 の WindowActivate downstream observation contract 固定は完了済み。
6. R08 の pending retry owner file boundary separation は完了済み。
7. R09 の window resolver / `activateWorkbook` route matrix は docs freeze 済み。Phase 4 では runtime extraction しない。
8. R13 foreground outcome contract は docs freeze 済み。foreground linkage / completion 接続は Phase 5 候補として残し、Phase 4 では runtime extraction しない。
9. R04/R14 display session は Phase 5 候補として残す。

## Phase 4 closure note

Phase 4 safe-first ownership separation は、ここで終了扱いにします。

Phase 4 は「巨大クラスを小さくする」フェーズではありませんでした。目的は、route / retry / trace / completion の意味を変えずに、safe-first で切れる owner を分離し、lifecycle visibility と danger boundary localization を高め、Phase 5 前に freeze line を安定させることでした。

完了扱いにする runtime separation:

| unit | 完了内容 | safe-first として扱えた理由 |
| --- | --- | --- |
| R02 refresh precondition / fail-closed policy | `TaskPaneRefreshPreconditionPolicy` に policy boundary を分離。 | completion / retry / foreground owner ではなく、fail-closed 判定の局所化だったため。 |
| R16 timer lifecycle ownership | `TaskPaneRetryTimerLifecycle` に timer create / register / stop / unregister / dispose を集約。 | retry sequencing と callback meaning を変えず、lifecycle visibility だけを高めたため。 |
| R06 ready-show retry scheduler ownership | `TaskPaneReadyShowRetryScheduler` に 80ms retry schedule / firing observation を分離。 | attempt sequencing と fallback handoff を既存 owner に残したため。 |
| R10/R11/R12 normalized outcome mapping | visibility / source / rebuild fallback の normalized outcome mapping を固定。 | raw result を completion に読み替えず、decision object の意味だけを局所化したため。 |
| R15 WindowActivate downstream observation | `WindowActivateDownstreamObservation` に downstream observation を分離。 | `WindowActivate dispatch != completion` を維持し、trace / route observation に限定したため。 |
| R08 pending retry owner file boundary | `PendingPaneRefreshRetryService` を pending retry owner file boundary として明確化。 | pending retry success / active fallback success を completion と読まない前提を維持したため。 |

Phase 4 終了前に docs freeze 済みの protocol boundary:

- ready-show retry contract truth。
- R07 pending fallback semantics。
- R09 window resolver / `activateWorkbook` route matrix。
- active fallback truth table。
- foreground outcome contract。

残っている候補は、もはや safe-first runtime extraction ではなく display recovery protocol の核心領域です。

| area | Phase 5 送りにする理由 | Phase 4 runtime extraction STOP |
| --- | --- | --- |
| R04/R14 display session | display-handoff と `case-display-completed` one-time emit の session boundary を持つ。 | state bag 化すると completion owner が見えなくなるため STOP。 |
| R05 callback/completion convergence | worker raw result を normalized outcomes と completion 判定へ再収束させる橋。 | callback meaning を変えやすいため STOP。 |
| R07 fallback handoff | ready-show exhaustion、immediate refresh、WorkbookOpen stabilization、pending retry entry をまたぐ。 | retry sequencing / fallback meaning に触れるため STOP。 |
| R09 window resolver | route-specific `activateWorkbook` intent と window availability boundary を持つ。 | foreground / completion / retry success と誤結合しやすいため STOP。 |
| R13 foreground linkage | `RequiredSucceeded` / `RequiredDegraded` / `NotRequired` を display-completable terminal として completion input にする。 | foreground outcome semantics を丸めやすいため STOP。 |

これ以上 Phase 4 で runtime extraction を続けると、completion semantics、callback meaning、retry sequencing、display session boundary、foreground outcome semantics、trace contract に近づきます。その作業は safe-first ownership separation ではなく protocol rewrite になりやすいため、Phase 4 の scope から外します。

Phase 5 の入口は、単なる orchestration shrink ではなく protocol-preserving convergence redesign です。Phase 5 では次を同じ設計面として扱います。

- completion ownership clarification。
- callback meaning clarification。
- display session convergence mapping。
- foreground linkage と completion convergence の距離固定。
- retry convergence と display session completion の接続整理。

Phase 5 初手候補:

- R05 callback/completion convergence。
- R10/R13/R14 foreground + completion convergence。
- display protocol convergence map。

Phase 4 closure 後も immutable freeze line として維持するもの:

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

## Phase 3 へ渡す freeze line 候補

Phase 3 では、下記候補を `docs/taskpane-display-recovery-freeze-line.md` の変更禁止契約として固定します。

- `WorkbookOpen` window-dependent skip 条件。
- `ReleaseWorkbook -> EnsureVisible -> SuppressUpcomingCasePaneActivationRefresh -> ShowWorkbookTaskPaneWhenReady` の順序。
- ready-show `attempt 1 -> 80ms retry attempt 2 -> pending retry fallback` の順序。
- pending retry `400ms / 3 attempts` と active CASE context fallback。
- `ResolveWorkbookPaneWindow(..., activateWorkbook: true/false)` の route 別呼び分け。
- WindowActivate gate の `case protection -> external workbook detection -> case pane suppression -> refresh dispatch`。
- normalized outcome の terminal / display-completable の意味。
- `RequiredDegraded` を display failure / success へ読み替えないこと。
- `WindowActivateDispatchOutcome.Dispatched` を display completion とみなさないこと。
- `case-display-completed` の emit owner と one-time completion 条件。

## 今回行わないこと

- コード変更。
- service 新設。
- helper 抽出。
- class rename。
- namespace 移動。
- retry 値 / 順序変更。
- trace 名変更。
- route contract 変更。
- fail-closed 条件変更。
- COM restore 順序変更。
- UI policy 変更。
- build / test / `DeployDebugAddIn` 実行。
