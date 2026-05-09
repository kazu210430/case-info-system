# White Excel Prevention Boundary Current State

## 位置づけ

この文書は、hidden Excel / isolated app / white Excel lifecycle redesign の G-0 として、white Excel recovery / prevention の current-state と target boundary を docs-only で整理するためのものです。

- 開始時の `main` / `origin/main`: `3311b5e6d0bbd7e07bc05f74f5646f139fdc3292`
- 作業ブランチ: `codex/white-excel-prevention-boundary-current-state`
- 対象: white Excel prevention / recovery の owner、trigger、primitive、trace、post-close follow-up、visible workbook 判定、target workbook still-open 判定、隣接 protocol との境界
- 非対象: コード変更、tests 変更、Quit 条件変更、visible workbook 判定変更、post-close follow-up 条件変更、WorkbookClose / reopen 条件変更、visibility / foreground 条件変更、hidden cleanup / retained cleanup 条件変更、WindowActivate dispatch 変更、build / test / `DeployDebugAddIn`

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
- `docs/case-workbook-lifecycle-current-state.md`
- `docs/case-display-recovery-protocol-current-state.md`
- `docs/case-display-recovery-protocol-target-state.md`

読取確認した主な実装 / tests:

- `PostCloseFollowUpScheduler`
- `CaseWorkbookLifecycleService`
- `CaseWorkbookBeforeClosePolicy`
- `CaseWorkbookOpenStrategy`
- `KernelCasePresentationService`
- `WorkbookWindowVisibilityService`
- `ExcelWindowRecoveryService`
- `TaskPaneRefreshOrchestrationService`
- `WindowActivatePaneHandlingService`
- `PostCloseFollowUpSchedulerTests`
- `CaseWorkbookLifecycleServiceThinOrchestrationTests`
- `CaseWorkbookOpenStrategyTests`
- `WindowActivatePaneHandlingServiceTests`

この文書は current-state を棚卸しするだけです。不明点は補完せず、G-1 以降の確認事項として残します。

## H consolidation note

この文書は white Excel prevention の detail current-state です。top-level lifecycle の正本は `docs/hidden-excel-isolated-app-white-excel-lifecycle-current-state.md`、outcome / owner vocabulary の正本は `docs/hidden-excel-lifecycle-outcome-vocabulary.md` として読む。

- current runtime で実装済みとして読む語は `white Excel prevention` です。
- `white Excel recovery` は、すでに白 Excel になった状態への user-facing UX / manual guidance / retry protocol を指す未定義領域です。`ExcelWindowRecoveryService` の CASE display recovery primitive と混同しません。
- `WhiteExcelPreventionNotRequired` は current-state では `targetWorkbookStillOpen` と `visibleWorkbookExists` の 2 つの主な reason を持ちます。どちらも quit しない decision であり、display success ではありません。
- `WhiteExcelPreventionSkipped` は vocabulary / target boundary 上の候補であり、現行 primary emitted outcome としては扱いません。
- post-close follow-up は queued key と current `Application.Workbooks` の fresh facts を比較します。follow-up cancel / reopen gating は current-state では未定義です。

## current-state summary

- white Excel prevention は close / quit protocol です。表示回復、foreground guarantee、WindowActivate dispatch、hidden cleanup、retained cleanup の代替ではありません。
- 現行 owner は `PostCloseFollowUpScheduler` です。`CaseWorkbookLifecycleService` は close 前に workbook key / folder path を確定して follow-up を予約する upstream owner です。
- follow-up は close 済み workbook COM object を再参照しません。queued workbook key と current `Application.Workbooks` の fresh enumeration から still-open / visible workbook を判断します。
- `targetWorkbookStillOpen` は queued key と current open workbook key が一致したことを表します。この場合は `WhiteExcelPreventionNotRequired` として quit しません。
- `visibleWorkbookExists` は target workbook が closed と読めた後、current shared app に visible window を持つ workbook が 1 冊以上あることを表します。この場合も `WhiteExcelPreventionNotRequired` として quit しません。
- still-open が false かつ visible workbook が無い場合だけ、`PostCloseFollowUpScheduler` が `DisplayAlerts=false` にして `Application.Quit()` を試みます。
- quit 成功時は終了中 `Application` の `DisplayAlerts` を restore しません。quit 失敗時だけ snapshot がある場合に restore します。
- 現行 emitted outcome は主に `WhiteExcelPreventionQueued` / `WhiteExcelPreventionNotRequired` / `WhiteExcelPreventionCompleted` / `WhiteExcelPreventionFailed` です。`Skipped` は vocabulary / target boundary 上の候補ですが、現行 primary emitted outcome ではありません。
- immediate reopen 近接時は follow-up cancel ではなく、dequeue 時の fresh facts による still-open skip または visible workbook skip として読めます。
- WindowActivate は white Excel recovery / prevention owner ではありません。dispatch outcome は diagnostic trigger であり、quit 判定や recovery completion ではありません。

## white Excel prevention owner / trigger / primitive / trace owner

| protocol unit | owner | trigger | primitive / action | trace owner |
| --- | --- | --- | --- | --- |
| close policy / close前 facts capture | `CaseWorkbookLifecycleService` / `CaseWorkbookBeforeClosePolicy` | `WorkbookBeforeClose` | CASE / Base 判定、managed close 判定、dirty 判定、folder path 解決、follow-up 予約 | `CaseWorkbookLifecycleService` |
| post-close follow-up queue | `PostCloseFollowUpScheduler` | `Schedule(workbookKey, folderPath)` | `PostCloseFollowUpRequest` enqueue、UI dispatcher / retry timer | `PostCloseFollowUpScheduler` |
| target still-open check | `PostCloseFollowUpScheduler` | queued request dequeue | queued workbook key と current `Application.Workbooks` enumeration から得た fresh open workbook key の比較 | `post-close-follow-up-decision` |
| visible workbook check | `PostCloseFollowUpScheduler` | target still-open が false | current `Application.Workbooks` を列挙し、`openWorkbook.Windows.Count > 0` かつ `window.Visible` を scan | `white-excel-prevention-outcome` |
| no visible workbook quit | `PostCloseFollowUpScheduler` | visible workbook が無い | `DisplayAlerts` snapshot、`DisplayAlerts=false`、`Application.Quit()` | `white-excel-prevention-outcome` |
| Excel busy retry | `PostCloseFollowUpScheduler` | COM `0x800AC472` かつ attempts remaining | request を next attempt として requeue、timer retry | scheduler logs |
| quit failure restore | `PostCloseFollowUpScheduler` | `Application.Quit()` exception | `WhiteExcelPreventionFailed` を記録し、snapshot がある場合だけ `DisplayAlerts` restore | `white-excel-prevention-outcome` |

## white Excel recovery と prevention の違い

current-state では、white Excel 対策の runtime behavior は prevention として実装されています。

- prevention:
  - CASE close 後に visible workbook が無い shared/current app を残さないため、post-close follow-up で `Quit` を試みることです。
  - owner は `PostCloseFollowUpScheduler` です。
  - trigger は close 前に予約された follow-up request です。
  - success は `WhiteExcelPreventionCompleted` です。
- recovery:
  - すでに白 Excel になった状態をユーザー向けに復旧する UX / manual guidance / retry protocol は current-state では未定義です。
  - `ExcelWindowRecoveryService` の visibility / foreground recovery は CASE display / refresh protocol の primitive であり、white Excel recovery owner ではありません。
  - `WindowActivate` は recovery trigger ではなく TaskPane display / refresh trigger です。

target boundary では、G-1 で outcome / trace を整理する場合も、まず existing prevention protocol の正規化に限定します。white Excel recovery の user-facing UX は別安全単位で仕様確認が必要です。

## post-close follow-up との境界

white Excel prevention は post-close follow-up の内部 protocol です。ただし post-close follow-up 全体と同義ではありません。

- post-close follow-up request は close 前に `CaseWorkbookLifecycleService` から予約されます。
- request payload は workbook key、folder path、attempt count です。
- folder path は payload に残りますが、white Excel prevention の visible workbook / quit primitive では使われません。
- queue dequeue 後に target still-open check を行い、still-open なら quit しません。
- still-open でない場合だけ visible workbook scan へ進みます。
- visible workbook がある場合は quit しません。
- visible workbook が無い場合だけ quit を試みます。

post-close follow-up 条件、retry count、retry interval、still-open check、visible workbook scan、quit 条件は G-0 では変更しません。

## visible workbook 判定の current meaning

current-state の visible workbook 判定は `PostCloseFollowUpScheduler.QuitExcelIfNoVisibleWorkbook()` 内の local fact です。

- 判定対象は current shared app の `_application.Workbooks` enumeration です。
- 各 open workbook について `Windows.Count > 0` かつ `Windows.Cast<Excel.Window>().Any(window => window.Visible)` を満たす場合に visible workbook exists と扱います。
- scan 中の例外は「closing workbook may already be tearing down」として ignore され、scan 継続されます。
- visible workbook が 1 冊でもあれば `WhiteExcelPreventionNotRequired` / `outcomeReason=visibleWorkbookExists` となり、quit しません。

この判定は white Excel prevention の quit eligibility だけを表します。CASE display の pane visible、workbook window visibility ensure、foreground guarantee、WindowActivate observation とは別概念です。

## target workbook still-open の current meaning

current-state の `targetWorkbookStillOpen` は、close 前に queue へ積まれた target key と、dequeue 時点の current open workbook facts の比較です。

- queued key は `CaseWorkbookLifecycleService` が close 前に確定した workbook key です。
- dequeue 時点で `PostCloseFollowUpScheduler.IsWorkbookStillOpen(workbookKey)` が current `_application.Workbooks` を列挙します。
- 各 open workbook の key は `ExcelInteropService.GetWorkbookFullName(workbook)` を優先し、空なら workbook name を使います。
- queued key と fresh open workbook key が一致した場合、target workbook is still open と扱います。
- この判定は close 済み workbook COM object の再参照ではありません。
- immediate reopen により fresh workbook が queued key と一致した場合も、current-state では follow-up cancel ではなく still-open skip として `WhiteExcelPreventionNotRequired` を記録します。

`targetWorkbookStillOpen` は `visibleWorkbookExists` とは別の skip 理由です。still-open は target identity の fresh open facts、visible workbook exists は target が closed と読めた後の application-wide visible window facts です。

## Quit success / failure / skipped / not-required の current meaning

| outcome | current emitted meaning | raw facts | mutation |
| --- | --- | --- | --- |
| `WhiteExcelPreventionQueued` | close 後 follow-up が予約された。 | workbook key、folder path presence、attempts remaining、queue count。 | queue mutation のみ。quit なし。 |
| `WhiteExcelPreventionNotRequired` / `targetWorkbookStillOpen` | queued target が current open workbook として残っているため quit 不要。 | `targetWorkbookStillOpen=True`、`quitAttempted=False`。 | quit なし。 |
| `WhiteExcelPreventionNotRequired` / `visibleWorkbookExists` | target は closed と読めるが shared/current app に visible workbook が残るため quit 不要。 | `hasVisibleWorkbook=True`、`quitAttempted=False`。 | quit なし。 |
| `WhiteExcelPreventionCompleted` | visible workbook が無く、`Application.Quit()` が完了した。 | `hasVisibleWorkbook=False`、`quitAttempted=True`、`quitCompleted=True`。 | `DisplayAlerts=false` 後に `Quit`。成功後 restore なし。 |
| `WhiteExcelPreventionFailed` | visible workbook が無く quit を試みたが失敗した。 | `hasVisibleWorkbook=False`、`quitAttempted=True`、`quitCompleted=False`。 | snapshot があれば `DisplayAlerts` restore、例外再送出。 |
| `WhiteExcelPreventionSkipped` | vocabulary / target boundary 上の候補。 | 現行 primary emitted outcome としては未確認。 | G-0 では実装しない。 |

`Completed` は reopen、visibility restore、foreground guarantee、WindowActivate dispatch の成功を意味しません。`NotRequired` は success / failure ではなく、existing conditions により quit しない decision です。

## hidden cleanup / retained cleanup との境界

white Excel prevention は hidden cleanup / retained cleanup の代替ではありません。

- hidden create session cleanup:
  - owner は `CaseWorkbookOpenStrategy` の hidden session mechanics です。
  - dedicated route は workbook close、hidden `Application.Quit`、COM release まで扱います。
  - outcome は `hidden-excel-cleanup-outcome` / `isolatedAppOutcome` で読む必要があります。
- retained hidden app-cache:
  - owner は `CaseWorkbookOpenStrategy` です。
  - session close で healthy cached app が idle に戻る場合は `RetainedInstanceReturnedToIdle` であり、cached app cleanup completed ではありません。
  - timeout / poison / feature flag disabled / shutdown で cached app disposal が起きた場合だけ `retained-instance-cleanup-outcome` で quit / release facts を読みます。
- white Excel prevention:
  - owner は `PostCloseFollowUpScheduler` です。
  - shared/current app の no visible workbook quit を扱います。
  - retained cached app、isolated hidden app、managed hidden reflection session は扱いません。

`RetainedInstanceReturnedToIdle`、`RetainedInstanceCleanupCompleted`、`HiddenExcelCleanupCompleted` を `WhiteExcelPreventionCompleted` に読み替えません。逆も同様です。

## visibility restore / foreground guarantee との境界

white Excel prevention は visibility restore / foreground guarantee の代替ではありません。

- `WorkbookWindowVisibilityService.EnsureVisible(...)` は workbook window の lightweight visible ensure です。
- `ExcelWindowRecoveryService.TryRecoverWorkbookWindow(...)` は `ScreenUpdating`、window visible、window state、application visible、activation、foreground promotion の full primitive です。
- `TaskPaneRefreshOrchestrationService` は `VisibilityRecoveryOutcome` と `ForegroundGuaranteeOutcome` を normalized outcome として扱い、created CASE display completion の条件に使います。
- `ForegroundGuaranteeOutcome.RequiredSucceeded` でも shared/current app quit は完了していません。
- `WhiteExcelPreventionCompleted` でも CASE display / foreground guarantee は完了していません。

visibility restore / foreground guarantee は CASE display / refresh protocol の領域です。post-close white Excel prevention は close / quit protocol の領域です。

## WorkbookClose / reopen との境界

WorkbookClose / reopen と white Excel prevention は近接しますが、owner と facts が違います。

- WorkbookClose:
  - `CaseWorkbookLifecycleService` が close 前に workbook key / folder path / policy facts を capture します。
  - managed close path では `ManagedCloseState` scope 内で close し、close 後に workbook object を再参照しない境界を維持します。
  - clean close path では close を cancel せず、post-close follow-up を予約します。
- post-close follow-up:
  - closed workbook object ではなく queued key と current application fresh facts を使います。
- reopen:
  - interactive created CASE の reopen は `KernelCasePresentationService` / `CaseWorkbookOpenStrategy.OpenHiddenForCaseDisplay(...)` が fresh workbook object を shared/current app 上で取得します。
  - reopen 後の window / role / DocProperty / TaskPane facts は display / refresh protocol 側で再取得します。

current-state では immediate reopen と follow-up の正式な cancel / gating protocol は未定義です。G-0 では still-open / visible workbook skip として現状を読むだけに留めます。

## WindowActivate が関与してはいけない範囲

`WindowActivate` は white Excel recovery / prevention owner ではありません。current-state / target boundary では、次を固定します。

- `WindowActivate` は post-close follow-up queue を作らない。
- `WindowActivate` は target still-open check を行わない。
- `WindowActivate` は visible workbook scan を行わない。
- `WindowActivate` は `Application.Quit()` を試みない。
- `WindowActivate` は quit success / failure / not-required outcome を emit しない。
- `WindowActivateDispatchOutcome.Dispatched` を `WhiteExcelPreventionCompleted` または `WhiteExcelPreventionNotRequired` に読み替えない。
- `WindowActivate` 発火を visible workbook exists の代替事実にしない。
- `WindowActivate` 発火を foreground guarantee completed、visibility recovery completed、CASE display completed の直接条件にしない。

`WindowActivatePaneHandlingService` は event facts を `TaskPaneDisplayRequest.ForWindowActivate(...)` に変換し、`Observed` / `Ignored` / `Deferred` / `Dispatched` / `Failed` の diagnostic dispatch outcome を記録する boundary です。

## raw facts と normalized outcome の分離

white Excel prevention では、raw facts と normalized outcome を分けて読む必要があります。

| layer | examples | owner | success / failure として読まないこと |
| --- | --- | --- | --- |
| request facts | workbook key、folder path presence、attempts remaining、queue count | `CaseWorkbookLifecycleService` / `PostCloseFollowUpScheduler` | queue されたことを quit success と読まない。 |
| still-open facts | queued key、fresh open workbook key、`targetWorkbookStillOpen` | `PostCloseFollowUpScheduler` | still-open skip を follow-up cancel と読まない。 |
| visible workbook facts | `hasVisibleWorkbook`、window visible scan | `PostCloseFollowUpScheduler` | pane visible / foreground visible と混同しない。 |
| quit raw facts | `quitAttempted`、`quitCompleted`、DisplayAlerts snapshot | `PostCloseFollowUpScheduler` | quit attempted を completed と読まない。 |
| normalized outcome | `WhiteExcelPreventionQueued` / `NotRequired` / `Completed` / `Failed` | `PostCloseFollowUpScheduler` | lower-level visibility / WindowActivate / hidden cleanup outcome で補完しない。 |

## current-state 上の未定義ポイント

- `WhiteExcelPreventionSkipped` を runtime emitted outcome として使う条件は未定義です。
- white Excel recovery の user-facing UX、manual guidance、retry-after-failure owner は未定義です。
- quit failure 後にユーザーへ何を表示するか、または追加 retry するかは未定義です。
- immediate reopen 近接時の正式な follow-up cancel / reopen gating protocol は未定義です。
- scheduler request の `folderPath` が white Excel prevention primitive で意味を持つかは未定義です。
- user close と system close を event facts だけで分ける normalized vocabulary は未定義です。
- visible workbook scan で COM exception を ignore した場合の diagnostic severity は未定義です。
- orphaned `EXCEL.EXE` 検出、通知、強制終了の top-level owner は未定義です。
- `WhiteExcelPreventionNotRequired` を target-state で `SkippedVisibleWorkbookExists` のように分けるかは未定義です。

## G-1 で触るべき安全単位

G-1 で実装へ進む場合も、最初の安全単位は runtime 条件を変えない outcome / trace 整理に限定します。

1. `PostCloseFollowUpScheduler` の existing emitted outcome と raw facts の意味を tests / docs で固定する。
2. `WhiteExcelPreventionNotRequired` の reason を `targetWorkbookStillOpen` と `visibleWorkbookExists` で明確に assertion する。
3. `WhiteExcelPreventionCompleted` が no visible workbook quit completed だけを意味し、foreground / visibility / WindowActivate / hidden cleanup success ではないことを boundary assertion する。
4. `WhiteExcelPreventionFailed` が quit attempted failure だけを意味し、DisplayAlerts restore は failure path だけであることを assertion する。
5. raw facts と normalized outcome を trace vocabulary 上で分け、`quitAttempted` / `quitCompleted` / `hasVisibleWorkbook` / `targetWorkbookStillOpen` を success terminal にしない。
6. current emitted outcome に `Skipped` を導入する場合も、既存 condition と meaning を変えず、互換 trace / tests を先に整理する。
7. related docs の参照導線を保ち、WorkbookClose / reopen / visibility / foreground / hidden cleanup / WindowActivate との相互除外を維持する。

## G-1 で触ってはいけない危険単位

- `PostCloseFollowUpScheduler.Schedule(...)` の予約条件。
- retry count / retry interval / Excel busy retry 条件。
- `IsWorkbookStillOpen(...)` の key 比較条件。
- visible workbook scan 条件。
- visible workbook がある場合の quit 抑止条件。
- no visible workbook 時の `Application.Quit()` 実行条件。
- quit 成功後に終了中 app を restore しない境界。
- quit failure 時だけ `DisplayAlerts` restore する境界。
- dirty prompt / folder offer / managed close scheduling 条件。
- WorkbookClose / reopen 条件。
- foreground / visibility recovery 条件。
- hidden create session cleanup 条件。
- retained hidden app-cache cleanup 条件。
- WindowActivate dispatch / activation 条件。
- service / helper 抽出、大規模責務移動、context-less workbook 推測、暗黙 open。
- `Application.DoEvents()`、sleep、timing hack、追加 guard で failure を覆う設計。
- orphaned `EXCEL.EXE` の broad kill。

## 今回行わないこと

- コード変更なし。
- tests 変更なし。
- Quit 条件変更なし。
- visible workbook 判定変更なし。
- post-close follow-up 条件変更なし。
- WorkbookClose / reopen 条件変更なし。
- visibility / foreground 条件変更なし。
- hidden cleanup / retained cleanup 条件変更なし。
- WindowActivate dispatch 変更なし。
- service / helper 抽出なし。
- build / test / `DeployDebugAddIn` 実行なし。docs-only 指示のため実行しません。

## 一言まとめ

white Excel prevention は、CASE close 後の shared/current app に visible workbook が残っていない場合だけ `PostCloseFollowUpScheduler` が `Application.Quit()` を試みる close / quit protocol です。`targetWorkbookStillOpen`、`visibleWorkbookExists`、quit success / failure は raw facts と normalized outcome を分けて読みます。WindowActivate、visibility restore、foreground guarantee、hidden cleanup、retained cleanup、WorkbookClose / reopen のいずれも white Excel prevention owner ではありません。
