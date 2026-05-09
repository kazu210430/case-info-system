# Visibility Restore / Foreground Guarantee Boundary Current State

## 位置づけ

この文書は、hidden Excel / isolated app / white Excel lifecycle redesign の F-0 として、visibility restore と foreground guarantee の境界 current-state を docs-only で棚卸しするためのものです。

- 開始時の `main` / `origin/main` / `HEAD`: `f5e6d8d669629553a48c582934c263cd8a68298d`
- 作業ブランチ: `codex/visibility-foreground-boundary-current-state`
- 非対象: コード変更、tests 変更、build / test / `DeployDebugAddIn` 実行

参照した正本:

- `AGENTS.md`
- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/case-display-recovery-protocol-current-state.md`
- `docs/case-display-recovery-protocol-target-state.md`
- `docs/hidden-excel-isolated-app-white-excel-lifecycle-current-state.md`
- `docs/hidden-excel-isolated-app-white-excel-lifecycle-target-state.md`
- `docs/hidden-excel-lifecycle-outcome-vocabulary.md`
- `docs/workbook-close-reopen-protocol-current-mapping.md`

今回の文書は current-state を記録するだけです。foreground primitive、retry 条件、visibility 判定、`WindowActivate` dispatch、hidden cleanup、white Excel prevention、WorkbookClose / reopen、retained cleanup / isolated app lifetime の条件は変更しません。

## H consolidation note

この文書は visibility / foreground の detail current-state です。top-level lifecycle の正本は `docs/hidden-excel-isolated-app-white-excel-lifecycle-current-state.md`、語彙の正本は `docs/hidden-excel-lifecycle-outcome-vocabulary.md` として読む。

- `visibility restore` は docs vocabulary の umbrella term です。保存状態正規化、hidden-for-display preparation、lightweight ensure、full recovery primitive をまとめて説明する時に使います。
- `visibility recovery` / `VisibilityRecoveryOutcome` は current code / trace vocabulary です。`TaskPaneRefreshOrchestrationService` が normalized outcome として扱う current-state 名として残します。
- `visibility restore` と `visibility recovery` は別 owner を示すための語彙ではありません。どちらも foreground guarantee、hidden cleanup、retained cleanup、white Excel prevention の代替ではありません。
- `foreground guarantee` は `TaskPaneRefreshOrchestrationService` が decision / outcome / trace owner、`ExcelWindowRecoveryService` が execution primitive owner です。`TaskPaneRefreshCoordinator` は execution bridge と raw facts return に留めて読む。
- `WindowActivate` は trigger / dispatch boundary です。visibility recovery owner、foreground guarantee owner、CASE display completed owner、cleanup owner として読まない。

## current-state summary

- visibility restore は単一 owner の 1 protocol ではなく、保存状態正規化、hidden-for-display 準備、presentation 前の lightweight ensure、refresh 前後の full recovery facts 正規化に分かれています。
- foreground guarantee は、refresh / ready-show 後に foreground obligation を terminal outcome にする protocol です。decision / outcome / trace owner は `TaskPaneRefreshOrchestrationService`、execution bridge は `TaskPaneRefreshCoordinator`、execution primitive は `ExcelWindowRecoveryService` です。
- `WorkbookWindowVisibilityService.EnsureVisible(...)` は workbook window の `Visible` を扱う lightweight primitive です。foreground promotion や hidden cleanup は扱いません。
- `ExcelWindowRecoveryService.TryRecoverWorkbookWindow(...)` は full recovery primitive です。`ScreenUpdating`、window visibility、window state、application visibility、window activation、foreground promotion を既存条件内で扱います。
- `TaskPaneRefreshOrchestrationService` は `VisibilityRecoveryOutcome` と `ForegroundGuaranteeOutcome` を normalized outcome として持ち、created CASE display session の `case-display-completed` 条件に使います。
- `WindowActivate` は window-safe な TaskPane display / refresh trigger です。visibility recovery owner、foreground guarantee owner、hidden cleanup owner、white Excel prevention owner、CASE display completed owner ではありません。
- hidden cleanup / retained cleanup は `CaseWorkbookOpenStrategy` と hidden session owner の領域です。visibility restore や foreground guarantee の成功を hidden cleanup 成功へ読み替えません。
- white Excel prevention は `PostCloseFollowUpScheduler` の close / quit protocol です。foreground recovery や visibility restore の代替ではありません。
- WorkbookClose / reopen は close lifecycle と presentation/open strategy の別 protocol です。reopen 後の visibility / foreground は display / refresh protocol 側で扱います。

## visibility restore owner / trigger / primitive / trace owner

| protocol area | current owner | trigger | primitive / action | trace owner |
| --- | --- | --- | --- | --- |
| hidden create save normalization | `KernelCaseCreationService` | hidden create session の save 前。`NormalizeInteractiveWorkbookWindowStateBeforeSave(...)` / `NormalizeBatchWorkbookWindowStateBeforeSave(...)` | `ResolveOrCreateWorkbookWindowForSave(...)`、必要時の `workbook.Activate()` / `workbook.NewWindow()`、`window.Visible = true`、minimized 時の `WindowState = xlNormal` | `NewCaseVisibilityObservation` の `save-window-*` |
| managed hidden reflection save normalization | `KernelUserDataReflectionService` | 未 open Base / Accounting 反映の save 前 | owner-owned hidden workbook window を保存前に restore し、hidden window state を保存ファイルへ残さない | `hidden-excel-cleanup-outcome` など hidden reflection owner 側の trace |
| hidden-for-display preparation | `CaseWorkbookOpenStrategy.OpenHiddenForCaseDisplay(...)` | interactive created CASE を shared/current app へ reopen する時 | shared app state snapshot、`ScreenUpdating` / `EnableEvents` / `DisplayAlerts` false、`Workbooks.Open(...)`、opened workbook window hide、previous window restore、shared app state restore | `shared-display-state-applied` / `shared-display-state-restored`、hidden-for-display logs |
| presentation initial recovery | `KernelCasePresentationService.ShowCreatedCase(...)` | hidden-for-display reopen 後、ready-show request 前 | `WorkbookWindowVisibilityService.EnsureVisible(...)`、`ExcelWindowRecoveryService.TryRecoverWorkbookWindowWithoutShowing(..., bringToFront: false)` | `initial-recovery-completed`、`Workbook window visibility recovery evaluated`、Excel window recovery trace |
| ready-show pre-visibility | `WorkbookTaskPaneReadyShowAttemptWorker` | ready-show attempt 1 回目 | `WorkbookWindowVisibilityService.EnsureVisible(...)`。attempt 2 回目以降は current-state tests 上、再実行しない | `ready-show-attempt`、`TaskPane wait-ready pre-visibility evaluated`、`ready-show-attempt-result` |
| refresh pre-context recovery | `TaskPaneRefreshCoordinator` | `TryRefreshTaskPane(...)` で Kernel HOME が visible でなく、recovery service がある時 | `TryRecoverWorkbookWindowWithoutShowing(...)` または active workbook 版。context / pane window 確定前の調整 | coordinator raw logs、Excel window recovery trace、`preContextRecoveryAttempted` / `preContextRecoverySucceeded` |
| visibility normalized outcome | `TaskPaneRefreshOrchestrationService` | precondition skip、refresh dispatch result、ready-show attempt outcome を受けた後 | mutation は持たず、`WorkbookWindowVisibilityEnsureFacts` と `TaskPaneRefreshAttemptResult` から `VisibilityRecoveryOutcome` を作る | `visibility-recovery-decision` / `visibility-recovery-*`。詳細 emit は created CASE display reason に寄る |
| full recovery primitive | `ExcelWindowRecoveryService` | presentation / refresh / foreground owner からの呼び出し | `ScreenUpdating=true`、window resolve / recreate、window visible、window state restore、`Application.Visible=true`、`window.Activate()`、foreground promotion | `Excel window recovery evaluated`、`Excel window recovery mutation trace` |

## foreground guarantee owner / trigger / primitive / trace owner

| protocol area | current owner | trigger | primitive / action | trace owner |
| --- | --- | --- | --- | --- |
| foreground decision / normalized outcome | `TaskPaneRefreshOrchestrationService` | refresh / ready-show attempt outcome の visibility outcome 後 | `ForegroundGuaranteeOutcome` を `NotRequired` / `SkippedAlreadyVisible` / `SkippedNoKnownTarget` / `RequiredSucceeded` / `RequiredDegraded` / `RequiredFailed` / `Unknown` で正規化 | `foreground-recovery-decision` |
| already-visible terminal | `WorkbookTaskPaneReadyShowAttemptWorker` が raw fact を返し、`TaskPaneRefreshOrchestrationService` が protocol outcome として扱う | visible CASE pane が既に対象 workbook window に shown | refresh せず `TaskPaneRefreshAttemptResult.VisibleAlreadySatisfied()`。foreground は `SkippedAlreadyVisible` | `taskpane-already-visible`、foreground decision trace |
| execution bridge | `TaskPaneRefreshCoordinator.ExecuteFinalForegroundGuaranteeRecovery(...)` | `TaskPaneRefreshOrchestrationService` が required と判断した時 | workbook があれば `TryRecoverWorkbookWindow(..., bringToFront: true)`、なければ active workbook 版 | execution result raw facts は orchestration 側の final trace へ戻る |
| execution primitive | `ExcelWindowRecoveryService` | execution bridge からの呼び出し | full recovery primitive。foreground promotion は `SetForegroundWindow` / top-most hop など既存 primitive 内で行う | `Excel window recovery evaluated`、mutation trace |
| final foreground trace | `TaskPaneRefreshOrchestrationService` | execution bridge 呼び出し前後 | mutation ではなく protocol trace。`RequiredSucceeded` / `RequiredDegraded` へ正規化 | `final-foreground-guarantee-started` / `final-foreground-guarantee-completed` |
| post foreground protection | `TaskPaneRefreshCoordinator` | foreground guarantee execution 後 | CASE context / window がある場合に protection を開始 | `protection-decision`。foreground guarantee completion そのものではない |

## visibility restore が扱うもの / 扱わないもの

visibility restore が扱うもの:

- owner-owned workbook window を hidden / minimized から visible / normal に戻すこと。
- hidden create session 内の save 前 normalization。
- managed hidden reflection session 内の save 前 normalization。
- hidden-for-display 後、shared/current app の表示 handoff 前に previous window / shared app state を restore すること。
- CASE presentation / ready-show の前に workbook window を解決し、必要なら visible にすること。
- refresh path の pre-context recovery raw facts を normalized visibility outcome へ渡すこと。
- pane visible state と visibility recovery facts を `case-display-completed` の材料として扱うこと。

visibility restore が扱わないもの:

- foreground guarantee の required / not-required 判定。
- `SetForegroundWindow` 等の final foreground obligation の semantic owner。
- hidden session cleanup、isolated app release、retained hidden app-cache cleanup。
- white Excel prevention の no visible workbook quit。
- `WorkbookClose` 条件、reopen 条件、post-close follow-up 条件。
- `WindowActivate` dispatch の success 判定。
- `Application.DoEvents()`、sleep、timing hack、追加 guard による表示不安定の隠蔽。

## foreground guarantee が扱うもの / 扱わないもの

foreground guarantee が扱うもの:

- display / refresh protocol 内で、pane visible 後に foreground obligation が残るかどうかを判定すること。
- required な場合に `TaskPaneRefreshCoordinator` 経由で `ExcelWindowRecoveryService` の full recovery primitive を呼ぶこと。
- foreground execution result を `ForegroundGuaranteeOutcome` に正規化すること。
- `case-display-completed` の条件として、foreground outcome が terminal かつ display-completable かを提供すること。

foreground guarantee が扱わないもの:

- workbook window `Visible=true` の lightweight ensure 自体。
- hidden / isolated app cleanup。
- retained hidden app-cache の keep / poison / timeout / shutdown cleanup。
- white Excel prevention の `Application.Quit()`。
- `WindowActivate` observed / dispatched を foreground success と呼ぶこと。
- WorkbookClose 後の still-open 判定、visible workbook scan、reopen gating。
- retry 回数、ready-show timing、foreground primitive の実行条件変更。

## raw facts / normalized outcome / trace owner / action owner

| layer | raw facts | normalized outcome owner | trace owner | action owner |
| --- | --- | --- | --- | --- |
| workbook window ensure | `WorkbookWindowVisibilityEnsureResult`、`WorkbookWindowVisibilityEnsureFacts`、`AlreadyVisible` / `MadeVisible` / `WindowUnresolved` 等 | display / refresh protocol 上は `TaskPaneRefreshOrchestrationService` の `VisibilityRecoveryOutcome` | `WorkbookWindowVisibilityService`、created CASE では orchestration の `visibility-recovery-*` | `WorkbookWindowVisibilityService` |
| full Excel recovery | bool return、`ScreenUpdating` / window visible / state / app visible / activation / foreground mutation facts | visibility outcome または foreground outcome の caller owner | `ExcelWindowRecoveryService` | `ExcelWindowRecoveryService` |
| foreground guarantee | foreground window / context / service availability / execution result | `TaskPaneRefreshOrchestrationService` の `ForegroundGuaranteeOutcome` | `TaskPaneRefreshOrchestrationService` | execution primitive は `ExcelWindowRecoveryService`、bridge は `TaskPaneRefreshCoordinator` |
| WindowActivate | `WindowActivateTaskPaneTriggerFacts` | `WindowActivatePaneHandlingService` の `WindowActivateDispatchOutcome` | raw event は `ThisAddIn`、dispatch outcome は `WindowActivatePaneHandlingService` | dispatch は `WindowActivatePaneHandlingService`。recovery / foreground action は持たない |
| hidden cleanup | workbook/app/cache facts、`workbookCloseAttempted`、`appQuitAttempted`、`cacheReturnedToIdle` 等 | hidden cleanup owner。CASE create mechanics は `CaseWorkbookOpenStrategy` | `hidden-excel-cleanup-outcome`、`retained-instance-cleanup-outcome` | hidden session owner / cache owner |
| white Excel prevention | queued workbook key、current open workbook scan、visible workbook scan、quit attempted / completed | `PostCloseFollowUpScheduler` の `WhiteExcelPrevention*` | `white-excel-prevention-outcome` | `PostCloseFollowUpScheduler` |
| CASE display completed | pane visible、visibility outcome、foreground outcome、refresh source、rebuild fallback facts | `TaskPaneRefreshOrchestrationService` | `case-display-completed` | lower-level action owner ではなく orchestration owner |

raw facts は success / failure ではありません。normalized outcome は protocol owner が作るものです。trace を emit した owner が、必ずしも primitive action owner になるわけではありません。

## WindowActivate が関与してよい範囲 / 関与してはいけない範囲

関与してよい範囲:

- Excel `WindowActivate` event の raw capture。
- `TaskPaneDisplayRequest.ForWindowActivate(...)` の作成。
- case protection による `Ignored`、case pane suppression による `Deferred`、refresh entry への `Dispatched` の記録。
- window-dependent な TaskPane display / refresh の trigger として downstream へ渡ること。
- `WorkbookActivate` と同様に host reuse / show / render path へ到達すること。

関与してはいけない範囲:

- `WindowActivate` 発火だけで visibility recovery completed とみなすこと。
- `WindowActivate` 発火だけで foreground guarantee completed とみなすこと。
- `WindowActivateDispatchOutcome.Dispatched` を CASE display completed の直接材料にすること。
- `WindowActivate` dispatch のために activation 条件を広げること。
- hidden create session、hidden-for-display、managed hidden reflection session、retained hidden app-cache の cleanup を持つこと。
- white Excel prevention、post-close follow-up、visible workbook absence quit を持つこと。
- WorkbookClose 条件、reopen 条件、close 後 COM 再参照リスクを補正すること。

関連 tests 上も、`WindowActivatePaneHandlingServiceTests.Handle_WhenAllowed_DispatchesDisplayRefreshTriggerWithoutOwningRecovery` は dispatch が recovery / foreground / hidden Excel owner ではないことを直接確認しています。

## hidden cleanup / white Excel prevention との境界

hidden cleanup:

- CASE create hidden session の mechanics / cleanup trace owner は `CaseWorkbookOpenStrategy` です。
- retained hidden app-cache の return-to-idle、poison、timeout、feature flag disabled、shutdown cleanup も `CaseWorkbookOpenStrategy` の cache owner 境界です。
- managed hidden reflection session は `KernelUserDataReflectionService` の owner 境界です。
- `VisibilityRecoveryOutcome.Completed` や `ForegroundGuaranteeOutcome.RequiredSucceeded` を `HiddenExcelCleanupCompleted` や `RetainedInstanceCleanupCompleted` に読み替えません。

white Excel prevention:

- close / quit 側の owner は `PostCloseFollowUpScheduler` です。
- queued key の still-open check、visible workbook scan、no visible workbook 時の `Application.Quit()` は close 前に capture した key と current `Application.Workbooks` の fresh facts で行います。
- visible workbook がある場合は quit しません。
- quit 成功後は終了中 `Application` を restore しません。quit failure 時だけ `DisplayAlerts` restore を試みます。
- foreground recovery、visibility restore、`WindowActivate` dispatch を post-close quit の代替にしません。

## WorkbookClose / reopen との境界

- WorkbookClose event boundary は `ThisAddIn.Application_WorkbookBeforeClose(...)`、close orchestration は `WorkbookLifecycleCoordinator.OnWorkbookBeforeClose(...)` です。
- CASE close policy / dirty prompt / managed close / clean close scheduling は `CaseWorkbookLifecycleService` 側の責務です。
- post-close follow-up は close 前に capture した workbook key を使い、closed workbook object を再参照しません。
- interactive created CASE の reopen は hidden create session close 後に `KernelCasePresentationService` / `CaseWorkbookOpenStrategy.OpenHiddenForCaseDisplay(...)` が fresh workbook object を shared/current app 上で取得する境界です。
- reopen 後の workbook window visibility、recovery、foreground、TaskPane completion は display / refresh protocol 側です。WorkbookClose owner や post-close follow-up owner ではありません。
- close 前 facts と reopen 後 factsを混ぜません。reopen は closed workbook の延長ではなく fresh workbook / window / context / DocProperty / snapshot / pane visibility facts の取得から始まります。

## retained cleanup / isolated app lifetime との境界

- isolated app は生成した owner が close / quit / COM release まで閉じます。
- CASE create の dedicated hidden route は `CaseWorkbookOpenStrategy` の mechanics owner が `Application.Quit` / release を扱います。
- retained hidden app-cache は healthy なら session close 後に idle へ戻り、cached `Application` は timeout / poison / feature flag disabled / shutdown でだけ cleanup されます。
- `RetainedInstanceReturnedToIdle` は cached app cleanup completed ではありません。
- `RetainedInstancePoisoned` は reuse 禁止であり、quit / release 完了ではありません。
- visibility restore は retained cleanup の代替ではありません。
- foreground guarantee は isolated app release の代替ではありません。
- `WindowActivate`、foreground recovery、white Excel prevention は retained cached app を推測で破棄しません。

## foreground / visibility / refresh source / rebuild fallback の近接境界

- visibility outcome は pane visible と recovery facts の normalized view です。
- foreground outcome は foreground obligation の terminal view です。
- refresh source selection は snapshot / cache source の normalized view です。
- rebuild fallback は `TaskPaneSnapshotBuilderService` の snapshot acquisition subprotocol です。
- `case-display-completed` は、created CASE display session に対して pane visible、visibility outcome display-completable、foreground outcome terminal / display-completable が揃った時だけ `TaskPaneRefreshOrchestrationService` が emit します。
- rebuild fallback や refresh source の成否を foreground success や visibility success と同一視しません。

## related tests current mapping

この F-0 では tests を実行しません。読取ベースで current mapping に関係する coverage は次の通りです。

- `WorkbookWindowVisibilityServiceTests`: already visible、hidden window の visible 化、window unresolved。
- `ExcelWindowRecoveryServiceTests`: missing workbook window の recreate と application / window visible / activation。
- `WorkbookTaskPaneReadyShowAttemptWorkerTests`: already-visible path は refresh せず `SkippedAlreadyVisible`、retry path は初回だけ visibility ensure。
- `WindowActivatePaneHandlingServiceTests`: allowed / protected / suppressed の dispatch outcome、recovery / foreground / hidden owner でないこと。
- `CaseWorkbookOpenStrategyTests`: hidden cleanup、retained cleanup、hidden-for-display shared app state restore。
- `PostCloseFollowUpSchedulerTests`: no visible workbook quit、quit failure restore、visible workbook skip、still-open skip。
- `TaskPaneManagerOrchestrationPolicyTests` / `TaskPaneHostReusePolicyTests`: WorkbookOpen window-dependent skip、WorkbookActivate / WindowActivate host reuse。
- `WorkbookCloseInteropHelperTests` / `ManagedCloseStateTests` / `CaseWorkbookLifecycleServiceThinOrchestrationTests`: close optional arguments、managed close scope、dirty / clean close scheduling、close pre-facts。

## current-state 上の未定義ポイント

- `VisibilityRecoveryOutcome` は current code / trace vocabulary、`VisibilityRestore*` は docs vocabulary / target-style naming です。current-state では `visibility restore` を umbrella term、`VisibilityRecoveryOutcome` を current emitted / observed outcome として併記します。
- `VisibilityRecoveryOutcome.Degraded` は current code 上 display-completable です。degraded を将来どの UX / trace severity とするかは未定義です。
- `ForegroundGuaranteeOutcome.RequiredFailed` は vocabulary と型にはありますが、current observed execution path は required execution後に `RequiredSucceeded` / `RequiredDegraded` へ寄っています。どの failure を `RequiredFailed` とするかは current-state では未定義です。
- `LogVisibilityRecoveryOutcome(...)` の detailed observation は created CASE display reason に寄ります。全 refresh reason に同じ normalized trace coverage を出すかは未定義です。
- `WorkbookActivate` と `WindowActivate` のどちらをすべての環境で最終安全境界にするべきかは未定義です。
- foreground guarantee が degraded / failed した場合の user-facing recovery guidance は未定義です。
- post-close quit failure 後の UX / retry / manual guidance は未定義です。
- orphaned retained hidden app-cache / orphaned `EXCEL.EXE` の検出、通知、強制終了 owner は未定義です。
- Kernel HOME visibility owner と CASE foreground guarantee owner が同時期に走る場合の protocol 名は未定義です。
- read-only / temporary workbook close 全般を hidden lifecycle 正本へ含めるかは未定義です。

不明な事項は current-state では補完しません。F-1 で条件変更や guard 追加に進む前に、別安全単位として確認が必要です。

## F-1 で触るべき安全単位

F-1 で触る場合も、最初の単位は runtime 条件を変えない観測 / vocabulary / tests の整理に限定します。

1. visibility restore と foreground guarantee の normalized outcome / trace の意味を tests または docs で固定する。
2. `WorkbookWindowVisibilityEnsureFacts`、`VisibilityRecoveryOutcome`、`ForegroundGuaranteeOutcome` の raw facts と normalized outcome の対応を追加 assertion で確認する。
3. `WindowActivateDispatchOutcome` が display success / recovery owner / foreground owner / hidden owner ではないことを維持する boundary assertion を広げる。
4. created CASE display session の `case-display-completed` が visibility outcome と foreground outcome の両方を見て成立することを trace / tests で確認する。
5. detailed trace coverage を増やす場合は `TaskPaneRefreshOrchestrationService` の outcome emit に限定し、primitive 条件は変えない。
6. docs 上の `visibility restore` / code 上の `visibility recovery` vocabulary を揃える場合も、`WorkbookWindowVisibilityService` / `ExcelWindowRecoveryService` の実行条件を変えない。
7. related tests は新規追加のみを安全単位とし、既存条件や fakes の意味を変更しない。

## F-1 で触ってはいけない危険単位

- `WorkbookWindowVisibilityService.EnsureVisible(...)` の visible 判定条件。
- `ExcelWindowRecoveryService` の `ensureWindowVisible`、`activateWindow`、`bringToFront`、`SetForegroundWindow`、`ShowWindow`、window state restore 条件。
- `TaskPaneRefreshOrchestrationService.CompleteForegroundGuaranteeOutcome(...)` の required / not-required 条件。
- `TaskPaneRefreshOrchestrationService.CompleteVisibilityRecoveryOutcome(...)` の display-completable 条件。
- ready-show retry 回数、attempt scheduling、first-attempt-only visibility ensure 条件。
- `WindowActivatePaneHandlingService.Handle(...)` の dispatch / ignore / defer 条件。
- `WorkbookOpen` 直後の window-dependent refresh skip 条件。
- hidden create session route、hidden-for-display、shared app state restore、hidden cleanup、retained cleanup 条件。
- `PostCloseFollowUpScheduler` の still-open check、visible workbook scan、quit / retry / `DisplayAlerts` restore 条件。
- WorkbookClose / reopen / dirty prompt / folder offer / managed close scheduling 条件。
- service / helper 抽出、大規模責務移動、context-less workbook 推測、暗黙 open。
- `Application.DoEvents()`、sleep、timing hack、追加 guard による表示 / foreground / cleanup failure の覆い隠し。

## 今回行わないこと

- コード変更なし。
- tests 変更なし。
- visibility restore 条件変更なし。
- foreground guarantee 条件変更なし。
- foreground primitive / retry 条件変更なし。
- `WindowActivate` dispatch 変更なし。
- hidden cleanup / white Excel prevention 条件変更なし。
- WorkbookClose / reopen 条件変更なし。
- retained cleanup / isolated app lifetime 条件変更なし。
- service / helper 抽出なし。
- build / test / `DeployDebugAddIn` 実行なし。docs-only 指示のため実行しません。

## 一言まとめ

current-state では、visibility restore は「対象 window を見える状態へ戻す、またはその raw facts を display protocol へ渡す」unit であり、foreground guarantee は「pane visible 後の foreground obligation を terminal outcome にする」unit です。どちらも `WindowActivate`、hidden cleanup、retained cleanup、white Excel prevention、WorkbookClose / reopen の代替 owner ではありません。
