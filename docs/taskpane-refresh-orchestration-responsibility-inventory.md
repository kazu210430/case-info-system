# TaskPaneRefreshOrchestrationService Responsibility Inventory

## 位置づけ

この文書は、TaskPane 表示回復領域の Phase 1 responsibility inventory です。

目的は `TaskPaneRefreshOrchestrationService` を今すぐ分割することではありません。現行の巨大 orchestration が何を背負っているかを安全単位で棚卸しし、Phase 2 で理想責務境界との対応表を作れる状態にします。

参照した正本:

- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/taskpane-display-recovery-current-state.md`
- `docs/taskpane-refresh-policy.md`
- `docs/visibility-foreground-boundary-current-state.md`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshOrchestrationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WorkbookTaskPaneReadyShowAttemptWorker.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshPreconditionPolicy.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WindowActivatePaneHandlingService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WorkbookLifecycleCoordinator.cs`

この文書は docs-only です。service 新設、class rename、namespace 移動、retry 順序変更、trace 名変更、route 契約変更、fail-closed 条件変更、UI policy 変更、COM restore 順序変更、実装リファクタ、メソッド移動は行いません。

## 全工程での位置

| Phase | 内容 | この文書との関係 |
| --- | --- | --- |
| Phase 0 | 正本 docs を固定する | 完了。`docs/taskpane-display-recovery-current-state.md` を入力にする。 |
| Phase 1 | responsibility inventory | 今回。現行 owner / 入出力 / retry / fail-closed / trace / 依存を棚卸しする。 |
| Phase 2 | 理想責務境界との対応表 | この文書の責務 ID を target boundary へ対応付ける。 |
| Phase 3 | 変更禁止条件固定 | `docs/taskpane-display-recovery-freeze-line.md` を正本として、この文書の「まだ分離してはいけない責務」を freeze line にする。 |
| Phase 4 | 安全単位で ownership 分離 | 責務 ID 単位で、順序と trace を変えずに owner を移せるか判断する。 |
| Phase 5 | orchestration 縮退 | Phase 4 後にだけ判断する。 |

## 対象範囲

対象:

- `TaskPaneRefreshOrchestrationService`
- `TaskPaneRefreshPreconditionPolicy.DecideRefreshPrecondition(...)`
- `RefreshDispatchShell`
- `WorkbookPaneWindowResolver`
- `PendingPaneRefreshRetryService`
- `WorkbookTaskPaneReadyShowAttemptWorker`
- `TaskPaneRefreshCoordinator`
- `TaskPaneRefreshPreconditionPolicy`
- `WindowActivatePaneHandlingService` から渡される display request / trace
- ready-show、retry、foreground outcome、window resolve、route、trace、fail-closed、owner 判定

対象外:

- code movement。
- service / helper の新設。
- `TaskPaneManager` / `TaskPaneHostFlowService` の host lifecycle 分離。
- hidden create / hidden-for-display / retained hidden app-cache の mechanics 変更。
- `KernelCasePresentationService` 本体の責務分離。
- `ThisAddIn` の VSTO adapter 分離。
- retry 値、trace 名、route contract、fail-closed 条件、UI policy、COM restore 順序の変更。

## 現在地

`TaskPaneRefreshOrchestrationService` は、現在の実装では「TaskPane を refresh するサービス」ではなく、TaskPane 表示回復 protocol の収束点です。

主に次を同時に持っています。

- refresh request の entry / normalized outcome / trace。
- ready-show request の受理。
- created CASE display session の開始と完了。
- ready-show retry timer。
- pending retry fallback の起動と timer cleanup。
- workbook window resolve helper。
- refresh precondition gate。
- `TaskPaneRefreshCoordinator` への dispatch。
- visibility / refresh source / rebuild fallback / foreground outcome の正規化。
- `WindowActivate` downstream trigger trace。
- `case-display-completed` の唯一 emit。

巨大化した理由は、無関係な処理を雑に足したためだけではありません。むしろ、window が安定しない Office event lifecycle、hidden-for-display 後の ready-show、pane already-visible path、foreground guarantee、protection / suppression、観測 trace を同じ protocol completion に束ねる必要があり、安全装置がこの class に集まったためです。

## 責務 Inventory

### R01. refresh entry / route normalization

| 項目 | 内容 |
| --- | --- |
| 現行 owner | `TaskPaneRefreshOrchestrationService.TryRefreshTaskPane(...)` |
| 入力 | raw `reason` または `TaskPaneDisplayRequest`、`Excel.Workbook`、`Excel.Window` |
| 出力 | `TaskPaneRefreshAttemptResult`、`try-refresh-start` / `try-refresh-end`、WindowActivate downstream trace |
| lifecycle 上の位置 | WorkbookActivate / WindowActivate / Startup / post-action / retry などから refresh path に入る最上流 |
| retry 有無 | 自身は retry しない。ready-show / pending retry から再入される |
| fail-closed 条件 | downstream precondition が refresh 不可なら skipped result へ丸め、success にはしない |
| trace 責務 | `try-refresh-start`、`try-refresh-end`、`window-activate-display-refresh-trigger-start`、`window-activate-display-refresh-trigger-outcome` |
| window dependency | 入力 window を保持し、必要な window resolve は downstream へ渡す |
| UI dependency | 直接 pane UI を作らない。UI 反映は coordinator / manager 側 |
| COM dependency | workbook/window descriptor と active state の観測に COM wrapper を使う |
| 変更理由 | route 追加、display request vocabulary 変更、trace field 追加 |
| 将来変更されやすい点 | raw string reason と structured request の併存、WindowActivate trace field |

### R02. refresh precondition gate

| 項目 | 内容 |
| --- | --- |
| 現行 owner | `TaskPaneRefreshPreconditionPolicy.DecideRefreshPrecondition(...)`。呼び出しと skip outcome への接続は `TaskPaneRefreshOrchestrationService.TryRefreshTaskPane(...)` |
| 入力 | `reason`、`workbook`、`window`、`ICasePaneHostBridge` |
| 出力 | proceed / `skip-workbook-open-window-dependent-refresh` / `ignore-during-protection` |
| lifecycle 上の位置 | refresh entry 直後、coordinator dispatch 前 |
| retry 有無 | なし。retry 中でも毎回同じ gate を通る |
| fail-closed 条件 | `WorkbookOpen` かつ workbook present かつ window null、または case protection 中 |
| trace 責務 | skip action 名は orchestration trace と outcome normalization の completion source になる |
| window dependency | `WorkbookOpen` skip では window null が条件 |
| UI dependency | なし |
| COM dependency | policy 自体は COM access なし。protection bridge 側は active window 判定を持つ |
| 変更理由 | fail-closed 条件整理、protection policy の owner 整理 |
| 将来変更されやすい点 | `ShouldIgnoreTaskPaneRefreshDuringCaseProtection(...)` が active window 基準で広めに止める点 |

### R03. refresh dispatch shell

| 項目 | 内容 |
| --- | --- |
| 現行 owner | `TaskPaneRefreshOrchestrationService.RefreshDispatchShell` |
| 入力 | `TaskPaneRefreshCoordinator`、`reason`、`workbook`、`window`、`KernelHomeForm` getter、suppression count getter |
| 出力 | `RefreshDispatchExecutionResult` / raw `TaskPaneRefreshAttemptResult` |
| lifecycle 上の位置 | precondition pass 後、outcome normalization 前 |
| retry 有無 | なし |
| fail-closed 条件 | coordinator が missing dependency / suppression / context reject を skipped or rejected として返す |
| trace 責務 | orchestration は dispatch result を `try-refresh-end` と normalized outcomes へ反映 |
| window dependency | coordinator 側で pane target window resolve が走ることがある |
| UI dependency | `TaskPaneManager.RefreshPaneWithOutcome(...)` へ到達しうる |
| COM dependency | coordinator の pre-context recovery / window resolve 経由 |
| 変更理由 | coordinator API 変更、suppression count / Kernel HOME visible condition 変更 |
| 将来変更されやすい点 | `KernelHomeForm` visibility が refresh pre-context recovery に効く点 |

### R04. ready-show request acceptance

| 項目 | 内容 |
| --- | --- |
| 現行 owner | `TaskPaneRefreshOrchestrationService.ShowWorkbookTaskPaneWhenReady(...)` |
| 入力 | created CASE workbook、ready-show reason |
| 出力 | ready-show enqueue trace、created-case display session、worker への `ShowWhenReady(...)` 委譲 |
| lifecycle 上の位置 | `KernelCasePresentationService` の post-release deferred presentation 後 |
| retry 有無 | worker / scheduler 経由で ready-show retry へ進む |
| fail-closed 条件 | worker は workbook null なら何もしない。session は created-case reason かつ workbook full name present の時だけ開始 |
| trace 責務 | `wait-ready-enqueued`、`ready-show-enqueued`、`created-case-display-session-started`、`display-handoff-completed` |
| window dependency | 入口では window は未確定。attempt 内で resolve |
| UI dependency | 直接 UI 表示しない。attempt が visible pane check / refresh delegate へ進む |
| COM dependency | workbook full name / active state の観測 |
| 変更理由 | created CASE presentation handoff、display session protocol、ready-show trace |
| 将来変更されやすい点 | ready-show をどの route に限定するか、created CASE 以外へ広げるか |

### R05. ready-show attempt delegation / completion callback

| 項目 | 内容 |
| --- | --- |
| 現行 owner | attempt 本体は `WorkbookTaskPaneReadyShowAttemptWorker`、callback owner は `TaskPaneRefreshOrchestrationService.HandleWorkbookTaskPaneShown(...)` |
| 入力 | `WorkbookTaskPaneReadyShowAttemptOutcome`、attempt number、workbook window、raw refresh result、visibility ensure facts |
| 出力 | visibility / refresh source / rebuild fallback / foreground outcome、completion check |
| lifecycle 上の位置 | ready-show attempt が shown と判定された直後 |
| retry 有無 | attempt failure 時は callback に来ず retry / fallback 側へ進む |
| fail-closed 条件 | outcome null なら終了。completion 条件未充足なら `case-display-completed` は出さない |
| trace 責務 | outcome normalization trace、completion trace |
| window dependency | attempt が解決した workbook window を消費 |
| UI dependency | visible CASE pane already shown facts を outcome として扱う |
| COM dependency | window descriptor / workbook full name 観測 |
| 変更理由 | ready-show attempt result vocabulary、display completion 条件 |
| 将来変更されやすい点 | already-visible path を display-completable と読む条件 |

### R06. ready-show retry timer

| 項目 | 内容 |
| --- | --- |
| 現行 owner | `TaskPaneRefreshOrchestrationService.ScheduleTaskPaneReadyRetry(...)` |
| 入力 | workbook、reason、attempt number、retry action |
| 出力 | `System.Windows.Forms.Timer`、80ms 後の retry action 実行 |
| lifecycle 上の位置 | ready-show attempt 1 が表示未成立だった後 |
| retry 有無 | ready-show retry。attempt 2 へ進める |
| fail-closed 条件 | retry action null なら timer を作らない |
| trace 責務 | `wait-ready-retry-scheduled`、plain `TaskPane wait-ready retry scheduled` |
| window dependency | 直接はなし。retry action 内で再 resolve |
| UI dependency | WinForms Timer を使う |
| COM dependency | timer scheduling 自体はなし。trace descriptor は workbook を読む |
| 変更理由 | retry delay / attempt logging / timer lifecycle |
| 将来変更されやすい点 | `80ms` が正式仕様値か経験値か未確定 |

### R07. pending retry fallback handoff

| 項目 | 内容 |
| --- | --- |
| 現行 owner | `TaskPaneRefreshOrchestrationService.ScheduleWorkbookTaskPaneRefresh(...)` |
| 入力 | workbook、reason |
| 出力 | workbook target tracking、immediate refresh attempt、pending retry sequence |
| lifecycle 上の位置 | ready-show attempts exhausted 後、または post-close / deferred workbook refresh |
| retry 有無 | pending retry に入る前に即時 refresh を 1 回試す。失敗時に 400ms retry |
| fail-closed 条件 | `WorkbookOpen` window-dependent skip なら fallback を開始しない |
| trace 責務 | `wait-ready-fallback-handoff`、`ready-show-fallback-handoff`、`defer-prepare`、`defer-immediate-success`、`defer-scheduled` |
| window dependency | fallback 開始前に `ResolveWorkbookPaneWindow(..., activateWorkbook: false)` |
| UI dependency | 直接 UI を作らない。refresh delegate が UI へ到達する |
| COM dependency | workbook full name tracking、window resolve |
| 変更理由 | fallback entry、workbook target / active target の扱い |
| 将来変更されやすい点 | immediate refresh と timer retry の境界、`WorkbookOpen` skip との関係 |

### R08. pending retry timer / state

| 項目 | 内容 |
| --- | --- |
| 現行 owner | `PendingPaneRefreshRetryService` |
| 入力 | tracked workbook full name or active target、reason、max attempts |
| 出力 | 400ms timer tick、workbook target refresh、active CASE context fallback refresh |
| lifecycle 上の位置 | ready-show fallback handoff 後、または explicit deferred active refresh |
| retry 有無 | pending retry max 3 attempts |
| fail-closed 条件 | attempts exhausted、target workbook unresolved かつ active context not CASE なら stop |
| trace 責務 | `defer-retry-start`、`defer-retry-end`、`defer-active-context-fallback-start`、`defer-active-context-fallback-end`、`defer-active-context-fallback-stop` |
| window dependency | workbook target retry では `ResolveWorkbookPaneWindow(..., activateWorkbook: true)` |
| UI dependency | WinForms Timer を使う |
| COM dependency | `FindOpenWorkbook(...)`、active context resolve、window resolve |
| 変更理由 | retry attempts、fallback target policy、timer cleanup |
| 将来変更されやすい点 | active CASE context fallback の必要条件と観測価値 |

### R09. workbook pane window resolve

| 項目 | 内容 |
| --- | --- |
| 現行 owner | `TaskPaneRefreshOrchestrationService.WorkbookPaneWindowResolver` |
| 入力 | workbook、reason、`activateWorkbook` |
| 出力 | first visible window、または active workbook matches + active window、または null |
| lifecycle 上の位置 | ready-show attempt、fallback prepare、pending retry、coordinator ensure-window |
| retry 有無 | synchronous resolve attempts 2 回 |
| fail-closed 条件 | interop service null、workbook null、visible/active window unresolved なら null |
| trace 責務 | `resolve-window-start`、`resolve-window-state`、`resolve-window-success`、`resolve-window-success-active-window`、`resolve-window-retry`、`resolve-window-failed` |
| window dependency | 中心責務。visible window / active workbook / active window を読む |
| UI dependency | なし。ただし `activateWorkbook=true` は Excel activation を伴う |
| COM dependency | `ActivateWorkbook(...)`、`GetFirstVisibleWindow(...)`、`GetActiveWorkbook(...)`、`GetActiveWindow(...)` |
| 変更理由 | window stability、activation policy、WorkbookOpen / WindowActivate 境界 |
| 将来変更されやすい点 | `activateWorkbook` をどの route で許すか、active window fallback の条件 |

### R10. visibility recovery outcome normalization

| 項目 | 内容 |
| --- | --- |
| 現行 owner | `TaskPaneRefreshOrchestrationService.CompleteVisibilityRecoveryOutcome(...)` |
| 入力 | reason、workbook/window、raw attempt result、completion source、attempt number、workbook window ensure facts |
| 出力 | `VisibilityRecoveryOutcome` attached result、visibility trace |
| lifecycle 上の位置 | precondition skip、refresh dispatch、ready-show callback 後 |
| retry 有無 | なし。retry result を消費する |
| fail-closed 条件 | raw facts が insufficient なら display completion へ進めない outcome になる |
| trace 責務 | `visibility-recovery-decision`、`visibility-recovery-*` |
| window dependency | pane visible / foreground window / input window / ensure facts を読む |
| UI dependency | pane visible facts を outcome として扱う |
| COM dependency | descriptor / observation correlation |
| 変更理由 | display-completable 条件、visibility vocabulary、trace coverage |
| 将来変更されやすい点 | degraded を display-completable とする意味、created CASE reason 以外への detailed trace 拡張 |

### R11. refresh source selection outcome normalization

| 項目 | 内容 |
| --- | --- |
| 現行 owner | `TaskPaneRefreshOrchestrationService.CompleteRefreshSourceSelectionOutcome(...)` |
| 入力 | raw attempt result、snapshot build result、reason、completion source |
| 出力 | `RefreshSourceSelectionOutcome` attached result、refresh source trace |
| lifecycle 上の位置 | visibility outcome 後、rebuild fallback outcome 前 |
| retry 有無 | なし |
| fail-closed 条件 | refresh source not reached / failed / fallback required の場合でも success に読み替えない |
| trace 責務 | `refresh-source-selected`、`refresh-source-degraded`、`refresh-source-fallback`、`refresh-source-rebuild-required`、`refresh-source-failed`、`refresh-source-not-reached`、`refresh-source-unknown` |
| window dependency | 間接。snapshot result は context/window に依存しうる |
| UI dependency | TaskPane snapshot / cache source の表示材料 |
| COM dependency | 直接 mutation なし |
| 変更理由 | snapshot source vocabulary、cache fallback、MasterList rebuild policy |
| 将来変更されやすい点 | raw `reason` を refreshSource として再掲している点 |

### R12. rebuild fallback outcome normalization

| 項目 | 内容 |
| --- | --- |
| 現行 owner | `TaskPaneRefreshOrchestrationService.CompleteRebuildFallbackOutcome(...)` |
| 入力 | raw attempt result、snapshot build result、refresh source outcome |
| 出力 | `RebuildFallbackOutcome` attached result、rebuild fallback trace |
| lifecycle 上の位置 | refresh source outcome 後、foreground outcome 前 |
| retry 有無 | なし |
| fail-closed 条件 | rebuild not reached / failed / cannot continue を display success へ丸めない |
| trace 責務 | `rebuild-fallback-required`、`rebuild-fallback-*` |
| window dependency | 間接 |
| UI dependency | snapshot acquisition / pane rendering facts |
| COM dependency | 直接 mutation なし |
| 変更理由 | MasterList rebuild fallback、snapshot/cache failure handling |
| 将来変更されやすい点 | rebuild fallback を display completion 条件に含めるかどうかの読み方 |

### R13. foreground guarantee decision / outcome

| 項目 | 内容 |
| --- | --- |
| 現行 owner | decision / trace は `TaskPaneRefreshOrchestrationService.CompleteForegroundGuaranteeOutcome(...)`、execution bridge は `TaskPaneRefreshCoordinator` |
| 入力 | raw attempt result、foreground context/window/workbook/service availability |
| 出力 | `ForegroundGuaranteeOutcome`、final foreground trace、post foreground protection |
| lifecycle 上の位置 | visibility / refresh source / rebuild fallback outcome 後、created-case completion 前 |
| retry 有無 | foreground 自体に retry はない |
| fail-closed 条件 | refresh not succeeded、pane not visible、refresh not completed、window null、recovery service null なら required execution しない |
| trace 責務 | `foreground-recovery-decision`、`final-foreground-guarantee-started`、`final-foreground-guarantee-completed` |
| window dependency | foreground window が required 判定の条件 |
| UI dependency | foreground / window recovery は UI-visible behavior に直結 |
| COM dependency | execution bridge 経由で `ExcelWindowRecoveryService` が workbook/window recovery と foreground promotion を実行 |
| 変更理由 | final foreground obligation、post foreground protection、display completion terminal condition |
| 将来変更されやすい点 | degraded の扱い、RequiredFailed の使い所、active workbook fallback |

### R14. created CASE display session / completion emit

| 項目 | 内容 |
| --- | --- |
| 現行 owner | `TaskPaneRefreshOrchestrationService.BeginCreatedCaseDisplaySession(...)` / `TryCompleteCreatedCaseDisplaySession(...)` |
| 入力 | created-case display reason、workbook full name、attempt result、visibility/foreground/refresh source/rebuild outcome |
| 出力 | session state、`case-display-completed`、`NewCaseVisibilityObservation.Complete(...)` |
| lifecycle 上の位置 | ready-show acceptance で session start、ready-show callback または refresh path 終端で completion check |
| retry 有無 | retry result を消費する。session 自体は retry しない |
| fail-closed 条件 | created-case reason でない、workbook missing、result missing、refresh not succeeded、pane not visible、visibility not terminal/display-completable、foreground not terminal/display-completable なら emit しない |
| trace 責務 | `created-case-display-session-started`、`display-handoff-completed`、`case-display-completed` |
| window dependency | completion details に window descriptor を含む |
| UI dependency | pane visible と foreground display-completable を completion 条件にする |
| COM dependency | workbook full name / descriptor |
| 変更理由 | created CASE display protocol、observability、final completion owner |
| 将来変更されやすい点 | completion 条件に refreshCompleted を必須化するか、already-visible path をどう扱うか |

### R15. WindowActivate downstream trigger trace

| 項目 | 内容 |
| --- | --- |
| 現行 owner | capture / dispatch は `WindowActivatePaneHandlingService`、downstream outcome trace は `TaskPaneRefreshOrchestrationService` |
| 入力 | `TaskPaneDisplayRequest.ForWindowActivate(...)`、workbook、window、refresh attempt result |
| 出力 | downstream trigger start / outcome trace |
| lifecycle 上の位置 | WindowActivate dispatch 後、refresh entry の開始と終了 |
| retry 有無 | なし。refresh path の結果を記録する |
| fail-closed 条件 | WindowActivate dispatched を display completion とみなさない |
| trace 責務 | `window-activate-display-refresh-trigger-start`、`window-activate-display-refresh-trigger-outcome` |
| window dependency | event window と downstream resolved facts を区別する |
| UI dependency | display request は UI refresh intent を含む |
| COM dependency | active state / workbook/window descriptor |
| 変更理由 | WindowActivate を recovery owner と誤読しないための observation |
| 将来変更されやすい点 | `TaskPaneDisplayRequest` と raw reason trace の二重管理 |

### R16. timer cleanup

| 項目 | 内容 |
| --- | --- |
| 現行 owner | `TaskPaneRetryTimerLifecycle`。停止入口は `TaskPaneRefreshOrchestrationService.StopPendingPaneRefreshTimer(...)`、pending retry callback owner は `PendingPaneRefreshRetryService` のまま |
| 入力 | pending retry timer start request、wait-ready retry timer request |
| 出力 | timer create / register / stop / unregister / disposal |
| lifecycle 上の位置 | immediate refresh success、pending retry success / exhausted / stop、ready-show shown callback、explicit stop |
| retry 有無 | retry を止める責務 |
| fail-closed 条件 | timer がない場合は何もしない |
| trace 責務 | 専用 trace は薄い。周辺の success / scheduled trace で観測する |
| window dependency | なし |
| UI dependency | WinForms Timer lifecycle |
| COM dependency | なし |
| 変更理由 | timer leak / duplicate retry 防止。R16 では retry semantics を動かさず、timer lifecycle owner だけを分離済み |
| 将来変更されやすい点 | R06 ready retry scheduler / R08 pending retry state の owner をさらに分けるか |

## 強く結合している責務

次は、現時点では変更理由が近い、または順序を崩すと挙動差分が出やすい責務です。

| 結合 | 理由 |
| --- | --- |
| R04 ready-show acceptance + R14 created CASE display session | `display-handoff-completed` と `case-display-completed` の同一 session を保つ必要がある。 |
| R05 ready-show callback + R10/R11/R12/R13 outcomes + R14 completion | completion は raw attempt result ではなく normalized outcomes を見て決まる。 |
| R06 ready retry + R07 fallback handoff + R08 pending retry | `attempt 1 -> 80ms attempt 2 -> 400ms pending fallback` の順序が安全装置。 |
| R02 precondition + R03 dispatch + R10/R11/R12 outcome normalization | skip / ignored でも normalized outcome と downstream trace が必要。 |
| R09 window resolve + R05 ready-show attempt + R08 pending retry | `activateWorkbook=true/false` の違いが route ごとの activation policy に関わる。 |
| R13 foreground outcome + R14 completion | foreground terminal / display-completable が `case-display-completed` の条件。 |
| R15 WindowActivate downstream trace + R01 route normalization | WindowActivate dispatch を display success と誤読させないため、start/outcome trace が refresh entry に隣接している。 |

## 一緒に動かすべき責務

Phase 4 で ownership 分離を検討する場合も、次は同じ安全単位として扱うべきです。

- R04 + R14: ready-show handoff と created CASE display session start/completion。
- R05 + R10 + R13 + R14: ready-show attempt outcome、visibility outcome、foreground outcome、final completion。
- R06 + R07 + R08: ready-show retry、fallback handoff、pending retry state。
- R02 + R01 skip trace: precondition skip と refresh entry trace / WindowActivate outcome trace。
- R09 activation policy + all callers: `activateWorkbook` の呼び分けと window resolve trace。

## 本来分離可能な責務

現時点で「すぐ切る」対象ではありませんが、責務としては分離候補にできるものです。

| 責務 | 分離可能に見える理由 | 分離前に必要な固定 |
| --- | --- | --- |
| R02 precondition gate | policy 判定は比較的 pure に近い | protection 判定の active window 基準を Phase 3 で凍結する |
| R06 ready retry timer | scheduling は attempt 本体と completion から分けられる | attempt 上限、80ms delay、timer cleanup owner |
| R08 pending retry state | 既に nested service として境界がある | active CASE context fallback の必要条件 |
| R09 window resolver | helper 化済みで入出力が明確 | `activateWorkbook` 許可 route、trace 名、attempt 数 |
| R10/R11/R12 outcome normalization | raw result から normalized outcome を作る責務 | display-completable 条件と created CASE detailed trace coverage |
| R15 WindowActivate downstream trace | event capture owner とは別の observation layer | display request fields と raw reason の契約 |
| trace formatting helpers | formatting は behavior と分けられる | trace 名と field vocabulary |

## まだ分離してはいけない責務

以下は、現在の理解では Phase 1 / Phase 2 で切ってはいけません。先に Phase 3 で freeze line を固定する必要があります。

| 責務 | まだ切ってはいけない理由 |
| --- | --- |
| R14 completion emit | `case-display-completed` が分散すると、already-visible path / refresh path / foreground terminal の意味が割れる。 |
| R13 foreground decision | execution primitive は別 owner だが、decision / outcome / completion condition は orchestration に寄っている。 |
| R04 ready-show acceptance と R14 session | session start と completion emit の相関が失われる。 |
| R05 ready-show callback と outcome normalization | raw attempt result を直接 completion に使う誤実装へ戻りやすい。 |
| R06/R07/R08 retry sequence | retry 順序そのものが flicker / unresolved window への安全装置。 |
| R09 `activateWorkbook` policy | route ごとの activation 副作用があり、切り出しだけでも挙動差分を生みやすい。 |
| R02 protection gate | active window 基準の広い TaskPaneRefresh protection が current-state 事実で、意図の正式文書は未確定。 |
| R15 WindowActivate downstream trace | `Dispatched` を completion と誤読しない observation contract がまだ重要。 |

## Orchestration が必要な理由

現行 orchestration が存在する理由は、複数の owner が返す raw facts を 1 つの display protocol outcome に収束させるためです。

- Excel event は `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` で window stability が異なる。
- created CASE 表示後は hidden-for-display、visibility ensure、suppression、ready-show が連鎖する。
- ready-show は already-visible path と refresh path の両方を成功相当に扱う。
- refresh execution は `TaskPaneRefreshCoordinator` / `TaskPaneManager` / host flow に分散している。
- foreground recovery は primitive owner と decision owner が異なる。
- `case-display-completed` は raw refresh success だけではなく、pane visible、visibility terminal、foreground terminal を見る。
- trace は user-visible failure investigation の contract になっている。

つまり orchestration は「全部を実行するため」ではなく、「順序、観測、completion 条件を 1 箇所で壊さないため」に必要です。

## Orchestration が危険化している理由

危険化している理由は、責務が多いことそのものではなく、変更理由が異なる責務が同じ class 内で近接しているためです。

- route 変更が trace contract と completion 条件に触れやすい。
- retry 値や fallback target の変更が ready-show completion に波及しやすい。
- window resolve の activation policy が foreground guarantee と混同されやすい。
- raw reason string と structured `TaskPaneDisplayRequest` が併存している。
- visibility / refresh source / rebuild fallback / foreground の normalized outcomes が同じ completion に集まる。
- `WindowActivate` は trigger でしかないのに、downstream trace が近いため recovery owner と誤読されやすい。
- protection / suppression は「止める」挙動として似て見えるが、lifecycle 上の意味が異なる。
- nested helper / nested service はあるが、変更禁止順序が class boundary と一致していない。

## 安全のために増えた責務

以下は行数削減だけを目的に削ると危険な、表示安定化のために増えた責務です。

- `WorkbookOpen` window-dependent refresh skip。
- ready-show retry `2 attempts / 80ms`。
- pending retry `400ms / 3 attempts`。
- active CASE context fallback。
- visible CASE pane already-visible early-complete。
- foreground guarantee decision / degraded outcome。
- post foreground protection。
- created CASE display session と one-time completion emit。
- WindowActivate dispatch is not display completion であることを示す trace。
- visibility / refresh source / rebuild fallback / foreground の normalized outcome。

## どこからなら安全に触れるか

Phase 1 時点ではコードを触りません。将来の安全単位候補としては、次の順で小さいです。

1. docs / tests で trace 名、route、retry 値、completion 条件を固定する。
2. pure policy に近い R02 の boundary tests は Phase 4 最初の safe unit として追加済み。
3. R09 window resolver の入出力と `activateWorkbook` contract を tests で固定する。
4. R10/R11/R12/R13 の normalized outcome mapping を tests で固定する。
5. R06/R08 の timer / retry state は、値と順序を固定した後に owner 分離を検討する。
6. R14 completion emit は最後に扱う。completion owner を移すなら Phase 3 の freeze line 後にする。

## 変更禁止として読む領域

- `WorkbookOpen` window-dependent skip 条件。
- ready-show `attempt 1 -> 80ms retry attempt 2 -> pending retry fallback`。
- pending retry `400ms / 3 attempts` と active CASE context fallback。
- `ReleaseWorkbook -> EnsureVisible -> SuppressUpcomingCasePaneActivationRefresh -> ShowWorkbookTaskPaneWhenReady`。
- WindowActivate gate の `case protection -> external workbook detection -> case pane suppression -> refresh dispatch`。
- foreground guarantee の required / not-required 条件。
- `case-display-completed` の emit owner と completion 条件。
- trace 名。
- route contract。
- fail-closed 条件。
- UI policy。
- COM restore / recovery primitive の実行順序。

## Phase 2 への入力

Phase 2 では、この inventory を次の target boundary へ対応付ける。

| Phase 1 responsibility | Phase 2 target boundary 候補 |
| --- | --- |
| R01 / R15 | display route / trigger observation boundary |
| R02 | refresh precondition / fail-closed policy boundary |
| R03 | refresh dispatch boundary |
| R04 / R14 | display protocol session boundary |
| R05 | ready-show attempt result boundary |
| R06 / R07 / R08 | retry / fallback ownership boundary |
| R09 | workbook pane window resolve boundary |
| R10 / R11 / R12 / R13 | normalized outcome boundary |
| R16 | timer lifecycle boundary |

Phase 2 でも、service を増やすことは成功条件にしません。変更理由単位で owner を見える化し、将来の変更コストを減らせる境界だけを候補にします。

Phase 2 の対応表は `docs/taskpane-refresh-orchestration-target-boundary-map.md` を正本とします。Phase 3 の freeze line は `docs/taskpane-display-recovery-freeze-line.md` を正本とします。

## 今回行わないこと

- コード変更。
- service 新設。
- class rename。
- namespace 移動。
- method move。
- retry 順序変更。
- trace 名変更。
- route 契約変更。
- fail-closed 条件変更。
- UI policy 変更。
- COM restore 順序変更。
- build / test / `DeployDebugAddIn` 実行。
