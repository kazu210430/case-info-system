# Lifecycle Redesign Cross Review

## 位置づけ

この文書は、hidden Excel / isolated app / retained hidden app-cache / white Excel lifecycle redesign の A-G 完了後レビューです。実装には入らず、現行 `main` の docs / tests / trace assertion を横断して、まだ残る ownership 混在、protocol 未定義、trace gap、次に触るべき安全単位を整理します。

H current-state consolidation 後は、この文書を A-G 完了時点の review / gap map として読む。current-state の source-of-truth / reference / detail docs の関係、語彙読み替え、current emitted outcome と target-only vocabulary の分離は、`docs/hidden-excel-isolated-app-white-excel-lifecycle-current-state.md` と `docs/hidden-excel-lifecycle-outcome-vocabulary.md` を正本にする。

- レビュー開始時の `main` / `origin/main` / `HEAD`: `e49ed38389cfae6a999f07689cceefb611ca9e10`
- 作業ブランチ: `codex/lifecycle-redesign-cross-review`
- 成果物: この docs-only 文書
- 非対象: コード変更、tests 変更、runtime 条件変更、build / test / `DeployDebugAddIn` 実行、service / helper 抽出、DoEvents / sleep / timing hack、guard 追加で覆う設計

参照した正本:

- `AGENTS.md`
- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/hidden-excel-isolated-app-white-excel-lifecycle-current-state.md`
- `docs/hidden-excel-isolated-app-white-excel-lifecycle-target-state.md`
- `docs/hidden-excel-lifecycle-outcome-vocabulary.md`
- `docs/workbook-close-reopen-protocol-current-mapping.md`
- `docs/visibility-foreground-boundary-current-state.md`
- `docs/white-excel-prevention-boundary-current-state.md`
- `docs/case-workbook-lifecycle-current-state.md`
- `docs/case-display-recovery-protocol-current-state.md`
- `docs/case-display-recovery-protocol-target-state.md`

読取確認した主な tests / trace assertion:

- `CaseWorkbookOpenStrategyTests`
- `KernelUserDataReflectionServiceTests`
- `PostCloseFollowUpSchedulerTests`
- `CaseWorkbookLifecycleServiceThinOrchestrationTests`
- `WindowActivatePaneHandlingServiceTests`
- `WorkbookTaskPaneReadyShowAttemptWorkerTests`
- `RebuildFallbackOutcomeTests`
- `WorkbookWindowVisibilityServiceTests`
- `ExcelWindowRecoveryServiceTests`

## A-G 完了後の到達点

A-G 後の到達点は、runtime 条件を大きく動かす前に、owner / outcome / trace の言葉を分けるところまで進んだ状態と読む。

| 単位 | 到達点 | まだ混ぜないもの |
| --- | --- | --- |
| A. lifecycle vocabulary | hidden cleanup、isolated app、retained cleanup、WorkbookClose / reopen、visibility、foreground、white Excel、WindowActivate の vocabulary が docs 上で定義された。 | vocabulary を理由に enum / helper / runtime 条件へ一気に広げない。 |
| B. hidden cleanup trace | `hidden-excel-cleanup-outcome` は `CaseWorkbookOpenStrategy` と `KernelUserDataReflectionService` 側で owner facts と outcome を出す。 | hidden workbook close 成功を cached app cleanup 成功へ読み替えない。 |
| C. isolated app lifetime | dedicated hidden route と managed hidden reflection session は生成 owner が close / quit / COM release まで閉じる境界として固定された。 | shared/current app quit 条件と混ぜない。 |
| D. retained app-cache | return-to-idle / poison / timeout / feature flag disabled / shutdown cleanup が `CaseWorkbookOpenStrategy` の cache owner 境界として整理された。 | `RetainedInstanceReturnedToIdle` を cleanup completed と呼ばない。 |
| E. WorkbookClose / reopen | close 前 immutable facts、post-close follow-up、immediate reopen 近接時の still-open / visible scan の current behavior が docs 化された。 | follow-up cancel / reopen gating / close 後 COM 再参照を追加しない。 |
| F. visibility / foreground | visibility restore と foreground guarantee を別 unit とし、decision / outcome / trace owner を `TaskPaneRefreshOrchestrationService` 側へ寄せる方針が固定された。 | visibility restore 成功を foreground success や hidden cleanup success と読まない。 |
| G. white Excel prevention | `PostCloseFollowUpScheduler` の no visible workbook quit protocol として `WhiteExcelPrevention*` outcome が整理された。 | WindowActivate、foreground、visibility、hidden cleanup を white Excel prevention owner にしない。 |

CASE display / recovery 側では、`case-display-completed` の final owner は `TaskPaneRefreshOrchestrationService` へ寄っており、`WindowActivate` は trigger / dispatch 境界として切り分けが進んでいる。一方、current-state docs の本文には historical section と最新到達点が混在しており、そこが次の docs-only 整理対象になる。

## 残存 Ownership 混在

| 混在箇所 | 現在の分担 | 残るリスク | 判断 |
| --- | --- | --- | --- |
| CASE hidden create session | session owner は `KernelCaseCreationService`、hidden open / close mechanics と retained cache owner は `CaseWorkbookOpenStrategy`。 | operation owner と cleanup owner の trace を 1 つの success に丸めやすい。 | GO: trace / tests の owner assertion。HOLD: route 条件変更。 |
| retained hidden app-cache | session close は hidden cleanup、cached app disposal は retained cleanup。 | return-to-idle と cleanup completed の読み替え。 | GO: outcome assertion 強化。HOLD: cache 撤去、timeout / feature flag 条件変更。 |
| hidden-for-display / display completion | `OpenHiddenForCaseDisplay(...)` は shared app reopen と一時 hidden / previous window restore。`case-display-completed` は orchestration。 | hidden-for-display を CASE display completed と誤読する。 | GO: docs / trace linkage。HOLD: open / visibility 条件変更。 |
| visibility restore | save normalization、hidden reflection normalization、workbook visible ensure、full recovery primitive、orchestration outcome に分裂。 | `VisibilityRecoveryOutcome.Completed` と foreground / cleanup の混同。 | GO: boundary assertion。HOLD: `EnsureVisible` / recovery 条件変更。 |
| foreground guarantee | decision / outcome / trace は `TaskPaneRefreshOrchestrationService`、execution primitive は `ExcelWindowRecoveryService`。 | one-shot promotion、WindowActivate、final guarantee の混同。 | GO: trace assertion。HOLD: foreground primitive 条件変更。 |
| WindowActivate | event capture は `ThisAddIn`、dispatch は `WindowActivatePaneHandlingService`、display entry は add-in boundary、refresh outcome は orchestration。 | trigger が recovery / foreground / cleanup owner に戻る。 | GO: diagnostic outcome assertion。HOLD: WindowActivate 挙動変更。 |
| WorkbookClose / reopen / post-close | close policy は `CaseWorkbookLifecycleService`、follow-up / quit は `PostCloseFollowUpScheduler`、reopen は `KernelCasePresentationService` / `CaseWorkbookOpenStrategy`。 | close 前 facts、fresh reopen facts、dequeue 時 fresh workbook facts が混ざる。 | GO: diagnostic trace assertion。HOLD: cancel / gating 実装。 |
| white Excel prevention | close / quit protocol owner は `PostCloseFollowUpScheduler`。 | visible workbook fact を pane visible / foreground visible と混同する。 | GO: outcome reason assertion は継続。HOLD: quit 条件変更。 |
| refresh source / snapshot source / rebuild fallback | source selection は snapshot builder、normalized source / rebuild outcome は orchestration。 | `reason`、`refreshSource`、snapshot source、rebuild fallback が同じ言葉に戻る。 | GO: docs / tests 整理。HOLD: source 採用順序変更。 |
| Kernel HOME / CASE foreground | Kernel HOME は `KernelHomeForm` / `KernelWorkbookDisplayService`、CASE foreground は display / refresh protocol。 | CASE 作成中に Kernel を前景へ戻す regression。 | HOLD: 実装前に protocol 名を docs-only で固定。 |

## Trace / Assertion Gap

| Gap | 現在確認できること | 弱い点 | 次の扱い |
| --- | --- | --- | --- |
| `case-display-completed` trace | 実装側では `TaskPaneRefreshOrchestrationService` が emit する。 | tests 側で `case-display-completed` / `display-handoff-completed` の trace 文字列 assertion は未確認。 | GO: trace assertion 追加候補。ただし今回は変更しない。 |
| visibility / foreground trace | `visibility-recovery-decision`、`foreground-recovery-decision`、`final-foreground-guarantee-*` は実装側に存在する。 | tests は outcome 型の分離を確認するが、normalized trace の emit 範囲までは薄い。 | GO: tests-only の境界 assertion 候補。 |
| rebuild / refresh source trace | `RebuildFallbackOutcomeTests` は outcome 分離を確認している。 | created CASE display 以外の normalized source / rebuild trace 範囲は docs 上も未定義。 | GO: docs-only taxonomy。HOLD: 全 refresh reason へ trace 拡大。 |
| hidden cleanup trace | `CaseWorkbookOpenStrategyTests` は isolated completed / failed、retained returned-to-idle、shutdown、timeout、poison を確認する。`KernelUserDataReflectionServiceTests` も hidden cleanup trace を見る。 | `KernelCaseCreationService` の session owner と mechanics owner の相関を 1 trace で見る assertion は薄い。open failure は `poisoned` assertion 中心で normalized details が弱い。 | GO: owner correlation assertion。HOLD: cleanup 条件変更。 |
| retained cleanup target outcomes | current emitted outcome は `Completed` / `Skipped` / `Degraded` 付近。 | target vocabulary の `NotRequired` / `Failed` / `OwnershipUnknown` は emitted outcome として未確認。 | HOLD: 実装導入。GO: docs で target-only と明記。 |
| white Excel outcome | `PostCloseFollowUpSchedulerTests` は queued、completed、failed、visible skip、still-open skip、`WhiteExcelPreventionSkipped` 未 emit を確認する。 | quit failure UX、COM exception ignore severity、`Skipped` 導入条件は未定義。 | GO: docs-only。HOLD: UX / retry / outcome semantic 変更。 |
| close/reopen facts | `workbook-close-immutable-facts-captured` と post-close decision trace は存在し、一部 tests で確認される。 | close 後 COM 再参照禁止を広く assertion する tests は限定的。reopen 後 fresh facts と queue facts の相互排他 assertion は未整備。 | GO: diagnostic tests。HOLD: follow-up cancel / gating。 |
| WindowActivate non-owner | `WindowActivatePaneHandlingServiceTests` は display completion / recovery / foreground / hidden owner でないことを確認する。 | downstream の visibility / foreground / rebuild / source outcome と trigger source の紐づけ assertion は薄い。 | GO: trace / diagnostic assertion。HOLD: dispatch 条件変更。 |
| 実機 trace | `case-display-current-state` は実機確認済み record を持つ。 | A-G 全体を current `e49ed...` 上で一括した実機 trace matrix は docs 上にない。 | GO: docs-only 観測表。HOLD: 今回は実機未実行。 |

## Protocol 未定義リスト

- WorkbookClose 直後に同一 CASE が即 reopen された場合の正式な follow-up cancel / reopen gating protocol。
- `PostCloseFollowUpScheduler` の quit failure 後の user-facing UX / retry / manual guidance。
- orphaned `EXCEL.EXE` の検出、通知、強制終了の top-level owner。
- retained hidden app-cache を今後も維持するか、運用上の必要性と監視方針。
- user close と system close を Excel event facts だけで分ける normalized vocabulary。
- scheduler request の `folderPath` が white Excel prevention primitive 上で意味を持つか。
- `WhiteExcelPreventionSkipped` を runtime emitted outcome として使う条件。
- docs 上の `visibility restore` と code 上の `VisibilityRecoveryOutcome` の命名差をどう扱うか。
- `RequiredFailed` / `Unknown` 系 outcome をどの execution path で emitted outcome にするか。
- `WorkbookActivate` と `WindowActivate` のどちらを最終安全境界と呼ぶか。
- `TaskPaneDisplayRequest.Source`、downstream `reason`、normalized refresh source、snapshot source の taxonomy。
- `BaseCacheFallback` を refresh source protocol 上どう呼ぶか。
- `MasterListRebuild` failure / degraded snapshot の UX 上の terminal 意味。
- one-shot foreground promotion と final foreground guarantee の protocol 関係。
- Kernel HOME visibility owner と CASE foreground guarantee owner が同時期に走る場合の protocol 名。
- read-only / temporary workbook close 全般を hidden lifecycle 正本へ含めるか。
- close failure path で workbook object を error reporting にどこまで使ってよいか。

## Current-State / Target-State 不整合

不整合は runtime の不具合断定ではなく、docs の読み方として整理する。

- `case-display-recovery-protocol-current-state.md` は冒頭に最新 implementation delta を持つ一方、本文の historical current-state section には foreground guarantee decision / terminal trace owner を `TaskPaneRefreshCoordinator` とする記述が残る。現行 target / code では decision / outcome / trace owner は `TaskPaneRefreshOrchestrationService` 側へ寄っているため、historical section と current `main` section を分ける必要がある。
- 各 lifecycle docs は phase ごとの開始 hash を持つが、A-G 後の current `main` hash `e49ed38389cfae6a999f07689cceefb611ca9e10` での横断 index はない。これは履歴としては正しいが、最新正本として読む時に迷いが出る。
- target-state vocabulary には `VisibilityRestore*`、code / current-state には `VisibilityRecoveryOutcome` がある。概念は近いが、同一名ではない。
- target-state outcome には target-only の名前が含まれる。例: `RetainedInstanceCleanupNotRequired`、`RetainedInstanceCleanupFailed`、`RetainedInstanceOwnershipUnknown`、`WhiteExcelPreventionSkipped`。current emitted outcome として読むと誤る。
- `refreshSource` という既存 log field は normalized source としては未確定であり、target-state の `RefreshSourceSelectionOutcome` と同一視しないほうがよい。
- `case-display-completed` は success-only terminal だが、関連する `pane visible`、`refresh completed`、`foreground guarantee`、`rebuild fallback` の trace 粒度が docs / tests / runtime observation で完全には横並びになっていない。

## 実装に入ってよい候補

ここでいう GO は、次フェーズで実装に入るなら安全単位として成立しやすいという意味です。今回この文書では実装しません。

| 候補 | 判断 | 触ってよい範囲 | build / test / 実機確認タイミング |
| --- | --- | --- | --- |
| hidden cleanup / retained cleanup trace assertion | GO | existing emitted trace の assertion 追加。route / app kind / owner / raw facts / normalized outcome の確認。 | tests 変更時は `build.ps1` 標準入口で Compile / test。runtime trace を変える場合のみ `DeployDebugAddIn` と実機確認。 |
| WindowActivate diagnostic outcome assertion | GO | `WindowActivateDispatchOutcome` が recovery / foreground / hidden owner でないこと、trigger source と downstream reason の分離を tests / trace で確認。 | tests-only なら Compile / test。runtime emit を変えるなら Deploy 後に実機ログ確認。 |
| visibility / foreground outcome boundary assertion | GO | `VisibilityRecoveryOutcome` と `ForegroundGuaranteeOutcome` の display-completable / terminal / degraded / failed の分離。 | Compile / test。実機は CASE 作成直後の ready-show / foreground trace を変える時点。 |
| WorkbookClose / post-close diagnostic trace assertion | GO | close 前 immutable facts、dequeue、still-open、visible workbook scan の assertion。 | Compile / test。Quit 条件に触れる場合だけ実機必須。 |
| refresh source / rebuild fallback outcome assertion | GO | `CaseCache` / `BaseCache` / `BaseCacheFallback` / `MasterListRebuild` / `NotReached` の outcome assertion。 | Compile / test。source 採用順序は触らない。 |

## まだ Docs-Only で整理すべき候補

| 候補 | 判断 | 理由 |
| --- | --- | --- |
| A-G cross index の追加 | GO | phase ごとの historical hash と current `e49ed...` の正本導線を分ける必要がある。 |
| case-display current-state の historical/current 分離 | GO | foreground guarantee owner の旧記述と最新 delta が同じ文書内に共存している。 |
| protocol 未定義リストの target-state 反映 | GO | immediate reopen、quit failure UX、orphaned instance owner などは実装前に言葉を決める必要がある。 |
| `visibility restore` / `visibility recovery` 命名の対応表 | GO | docs vocabulary と code vocabulary の読み替え事故を避ける。 |
| `reason` / `source` / `snapshot source` / `rebuild fallback` taxonomy | GO | trace assertion 前に何を success / diagnostic と読むかを固定する必要がある。 |
| real-machine log checklist の横断表 | GO | A-G の実機観測観点を 1 か所にまとめるだけなら docs-only で安全。 |

## 触ってはいけない危険単位

- `WindowActivate` を cleanup / recovery / foreground guarantee / CASE display completed owner に戻すこと。
- `WorkbookOpen` 直後の window-dependent refresh skip 条件を弱めること。
- `PostCloseFollowUpScheduler` の still-open / visible workbook / quit / retry 条件を変更すること。
- visible workbook が 1 冊でもある shared/current app を white Excel prevention 名目で quit すること。
- close 済み workbook COM object を再参照すること。
- immediate reopen 近接時の cancel / gating を guard 追加で先に実装すること。
- retained hidden app-cache の feature flag、idle timeout、bypass、poison、shutdown cleanup 条件を変えること。
- orphaned `EXCEL.EXE` を process name / PID だけで強制終了すること。
- `WorkbookWindowVisibilityService` / `ExcelWindowRecoveryService` / foreground primitive の条件を trace 整理名目で変えること。
- ready-show retry 回数、attempt scheduling、first-attempt-only visibility ensure 条件を変えること。
- source selection order、`CaseCache -> BaseCache -> BaseCacheFallback -> MasterListRebuild` の採用順序を変えること。
- service / helper 抽出、大規模責務移動、context-less workbook 推測、暗黙 open。
- `Application.DoEvents()`、sleep、timing hack、追加 guard で failure を覆う設計。

## 実機確認で見るべきログ観点

実機確認が必要になるのは、runtime trace emit、runtime 条件、visibility / foreground primitive、close / quit に触る時点です。docs-only と tests-only では実行しません。

見るべき主なログ:

- `Runtime execution observed` の `assemblySha256`。Compile 成功と runtime `Addins\` 反映成功を混同しない。
- `NewCaseVisibilityObservation` と `KernelFlickerTrace`。新規 CASE 作成直後から初回表示完了までを見る。
- `hidden-excel-cleanup-outcome`: route、application kind、hidden cleanup outcome、isolated app outcome、retained instance outcome。
- `retained-instance-cleanup-outcome`: cleanup reason、quit attempted / completed、application lifetime owner。
- `workbook-close-immutable-facts-captured`、`workbook-close-follow-up-facts-captured`、`post-close-follow-up-request-dequeued`、`post-close-follow-up-decision`。
- `WhiteExcelPreventionQueued` / `NotRequired` / `Completed` / `Failed` と `targetWorkbookStillOpen` / `hasVisibleWorkbook`。
- `created-case-display-session-started`、`display-handoff-completed`、`ready-show-requested`、`ready-show-enqueued`、`ready-show-attempt`、`ready-show-attempt-result`。
- `visibility-recovery-decision`、`foreground-recovery-decision`、`final-foreground-guarantee-started`、`final-foreground-guarantee-completed`、`case-display-completed`。
- `Task pane snapshot source=CaseCache|BaseCache|BaseCacheFallback|MasterListRebuild`、`rebuild-fallback-*`、`refresh-source-*` 相当の normalized trace。
- `WorkbookOpen -> WorkbookActivate -> WindowActivate` の順序と、その時点の target workbook / target window 解決可否。
- `Application.Visible`、`ScreenUpdating`、workbook window `Visible`、`WindowState` の復元有無。

## 推奨する次フェーズ

| 優先 | フェーズ候補 | 判断 | 内容 | build / test / 実機確認 |
| --- | --- | --- | --- | --- |
| 1 | H. docs current-state consolidation | GO | A-G cross index、historical/current section の分離、target-only outcome の明示、未定義 protocol の集約。 | docs-only なら build / test / 実機確認なし。 |
| 2 | I. boundary trace assertion pack | GO | hidden / retained、WindowActivate、visibility / foreground、white Excel、close/reopen、refresh source の tests-only assertion を小単位で追加。runtime 条件は変えない。 | `build.ps1` 標準入口で Compile / test。Deploy は不要。ただし runtime log emit を変えるなら Deploy と実機確認。 |
| 3 | J. unresolved protocol decision docs | HOLD for implementation / GO for docs | immediate reopen cancel/gating、quit failure UX、orphaned instance owner、Kernel HOME vs CASE foreground の protocol 名を決める。 | docs-only は build / test 不要。実装は仕様決定後。 |

次に実装へ進むなら、もっとも安全なのは I の tests-only assertion pack です。ただし、その前に H の docs 整理で current-state / target-state の読み方を揃えると、trace assertion の期待値がぶれにくい。

## Build / Test / 実機確認が必要になるタイミング

- docs-only 変更のみ: build / test / `DeployDebugAddIn` / 実機確認は実行しない。
- tests-only assertion 追加: `build.ps1` を標準入口として Compile / test を行う。`dotnet build .\dev\CaseInfoSystem.slnx` は標準確認コマンドにしない。
- trace emit field / outcome emit の runtime 変更: Compile / test 後、runtime `Addins\` 反映が必要なら `.\build.ps1 -Mode DeployDebugAddIn` を使う。Excel を完全終了し、`Runtime execution observed` の `assemblySha256` で実行 DLL を確認する。
- visibility / foreground / close / quit / hidden cleanup 条件に触る変更: Compile / test / DeployDebugAddIn / 実機確認が必要。A-G の禁止条件に抵触しない小単位に分ける。
- service / helper 抽出、大規模責務移動、timing workaround、guard 追加: 現時点では HOLD。

## 結論

A-G で owner と vocabulary の土台はかなり固まっている。残っている主問題は、runtime 条件そのものよりも、historical docs と current docs の混在、target-only outcome と emitted outcome の読み違い、そして trace assertion の粒度不足である。

次は docs-only の current-state consolidation を先に置き、その後に tests-only の boundary assertion を小さく入れるのが安全。WorkbookClose / reopen cancel、white Excel recovery UX、orphaned instance cleanup、WindowActivate owner 化、visibility / foreground 条件変更は、まだ HOLD として扱う。
