# TaskPane 表示回復 Freeze Line / 変更禁止契約

## 位置づけ

この文書は、TaskPane 表示回復領域の Phase 3「freeze line / 変更禁止条件固定」の正本です。

目的はドキュメントを増やすことではありません。今後 `TaskPaneRefreshOrchestrationService` を安全に ownership 分離するため、ready-show、pending retry、foreground outcome、display session、WindowActivate trace、`case-display-completed`、protection / fail-closed の契約を、次フェーズ以降の refactor 合格基準として使える粒度で固定します。

参照した正本:

- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/taskpane-display-recovery-current-state.md`
- `docs/taskpane-refresh-orchestration-responsibility-inventory.md`
- `docs/taskpane-refresh-orchestration-target-boundary-map.md`
- 補助参照: `docs/taskpane-refresh-policy.md`
- 補助参照: `docs/visibility-foreground-boundary-current-state.md`
- 補助参照: `docs/case-display-recovery-protocol-current-state.md`

この文書は docs-only です。コード移動、service 新設、class rename、namespace 移動、retry 順序変更、trace 名変更、route 契約変更、fail-closed 条件変更、foreground outcome 条件変更、completion 条件変更、UI policy 変更、COM restore 順序変更、実装 refactor は行いません。

## Freeze Line の読み方

- `freeze line` は「今後の refactor で変えてはいけない runtime 契約」です。
- owner を移す場合も、ここに書いた順序、条件、trace 名、completion 意味を変えないことを合格条件にします。
- docs に根拠がない仕様値の意味は、正式業務仕様とは断定しません。ただし現行 runtime 契約としては変更禁止にします。
- 「helper 化してよいか」は、処理が小さいかどうかではなく、completion / retry / trace / fail-closed の意味を保てるかで判断します。

## 1. Ready-Show Retry 順序

### 固定する順序

created CASE 表示後の ready-show は、次の順序で固定します。

1. `KernelCasePresentationService` が transient suppression を release する。
2. `WorkbookWindowVisibilityService.EnsureVisible(...)` による Workbook Window 可視化を再実行する。
3. `SuppressUpcomingCasePaneActivationRefresh(...)` を設定する。
4. `ShowWorkbookTaskPaneWhenReady(...)` を呼ぶ。
5. `TaskPaneRefreshOrchestrationService` が `ready-show-enqueued` を記録する。
6. `TaskPaneRefreshOrchestrationService` が created CASE display session を開始し、`created-case-display-session-started` と `display-handoff-completed` を記録する。
7. `WorkbookTaskPaneReadyShowAttemptWorker` が attempt 1 を実行する。
8. attempt 1 だけ `WorkbookWindowVisibilityService.EnsureVisible(...)` を実行する。
9. `ResolveWorkbookPaneWindow(..., activateWorkbook: true)` で pane 対象 window を解決する。
10. visible CASE pane が同じ workbook / window に既にある場合は `taskpane-already-visible` として refresh を skip し、already-visible path の success 相当 raw fact を返す。
11. already-visible でない場合だけ `TryRefreshTaskPane(...)` へ refresh を handoff する。
12. attempt が表示成立しない場合だけ、`80ms` 後に attempt 2 を schedule する。
13. attempt 2 も表示成立しない場合だけ、`ready-show-fallback-handoff` で pending retry fallback へ渡す。

### retry とみなすもの

- ready-show retry と呼んでよいのは、attempt 1 の後に `TaskPaneReadyShowRetryScheduler` が `80ms` で attempt 2 を発火させる経路だけです。
- ready-show の max attempts は `2` として扱います。
- pending retry `400ms / 3 attempts` は ready-show exhaustion 後の fallback retry であり、ready-show retry そのものではありません。
- `ResolveWorkbookPaneWindow(...)` 内の window resolve attempts は window resolve の内部試行であり、ready-show retry と同一視しません。

### retry してはいけないもの

- attempt 1 だけの pre-visibility ensure を、根拠なく attempt 2 以降へ広げません。
- `TryRefreshTaskPane(...)` の下位 refresh 本体を、ready-show retry 名目で追加再実行しません。
- already-visible path を retry 対象に戻しません。
- CASE 専用 visible pane early-complete を accounting / Kernel / external workbook へ広げません。
- foreground guarantee execution を ready-show retry として再試行しません。
- hidden create / hidden-for-display / workbook close / retained cleanup を ready-show retry に含めません。
- trace emit の欠落を埋めるために retry を増やしません。trace は観測契約であり、retry 条件ではありません。

### Phase 4 合格基準

- `attempt 1 -> 80ms retry attempt 2 -> pending retry fallback` の順が変わっていないこと。
- `ReleaseWorkbook -> EnsureVisible -> SuppressUpcomingCasePaneActivationRefresh -> ShowWorkbookTaskPaneWhenReady` の順が変わっていないこと。
- attempt worker は attempt 本体までを扱い、pending retry state を持たないこと。
- orchestration は ready-show acceptance、session、fallback handoff、completion への収束を失っていないこと。

## 2. Pending Retry 条件

### pending fallback に入る条件

pending fallback は、ready-show attempts が尽きた場合にだけ、ready-show 表示回復の fallback として入ります。

固定する入口:

- ready-show attempt 1 と attempt 2 が表示成立しない。
- `TaskPaneRefreshOrchestrationService.ScheduleWorkbookTaskPaneRefresh(...)` へ handoff する。
- workbook target を tracking する。
- timer 開始前に immediate refresh を 1 回試す。
- immediate refresh が success した場合は pending retry timer を開始しない。
- immediate refresh が success しない場合だけ、`PendingPaneRefreshRetryService` が `400ms` interval / `3` attempts で retry する。

### R07 handoff semantics freeze line

`ScheduleWorkbookTaskPaneRefresh(...)` は、単なる timer helper ではありません。R07 の freeze line では、次を同じ protocol entry の意味として固定します。

- ready-show fallback handoff trace。
- `WorkbookOpen` skip。
- workbook target tracking。
- window resolve。
- pre-timer immediate refresh。
- pending retry start decision。

現行 code 上、この entry は created CASE ready-show exhaustion 後の handoff と、`KernelHomeForm.OpenSheet.PostClose` の workbook-target delayed refresh entry の両方から使われます。この二重性を固定し、現時点では ready-show exhaustion 専用 owner として runtime extraction しません。

immediate refresh は pending timer 開始前の re-entry です。次の読み替えは禁止します。

- immediate refresh success = recovered。
- immediate refresh success = `case-display-completed`。
- immediate refresh success = display recovery completed。

pending retry success も completion ではありません。immediate refresh / pending retry のどちらから戻った場合でも、completion は `TaskPaneRefreshOrchestrationService.TryCompleteCreatedCaseDisplaySession(...)` 相当の existing completion chain が、pane visible、visibility terminal / display-completable、foreground terminal / display-completable を満たした場合だけ成立します。

`WorkbookOpen` skip は incidental guard ではなく、window stability boundary に関する runtime stabilization contract です。`ready-show-fallback-handoff` trace 後でも pending retry start とは限りません。`WorkbookOpen` 直後に workbook が存在し window が未解決な場合は、pending retry を開始せず後続イベントへ委ねます。この trace -> skip -> no pending start の順序を変えません。

`PendingPaneRefreshRetryService` 内の active CASE context fallback は、tracked workbook を見失った場合の target-lost resiliency fallback です。これは completion fallback ではなく、active CASE context fallback が success したことだけを `case-display-completed` とみなしません。

R07 は fallback handoff semantics、immediate refresh re-entry、`WorkbookOpen` precondition、pending retry entry、display session boundary、normalized outcome chain、completion owner、trace contract にまたがるため、現時点では orchestration ownership に残します。

固定する開始禁止:

- `WorkbookOpen` 直後の `workbook != null && window == null` による window-dependent skip を、pending fallback 開始条件にしません。
- ready-show exhaustion 前に pending fallback へ落としません。
- target workbook / active CASE context のどちらも追えない状態を completion へ丸めません。

### pending から復帰できる条件

pending retry からの復帰は、timer が動いたことではなく、同じ refresh / outcome / completion chain に戻れた場合にだけ成立します。

復帰できる条件:

- tracked workbook を見つけ、pane 対象 window を解決し、refresh path が success する。
- tracked workbook を見失っても、active context が CASE として解決でき、active refresh fallback が success する。
- created CASE display session が存在し、後述の `case-display-completed` 条件を満たす。

復帰できない条件:

- pending retry attempts exhausted。
- target workbook unresolved かつ active context が CASE ではない。
- context rejected / missing dependency / suppression / protection により refresh が fail-closed した。
- pane visible、visibility terminal、foreground terminal / display-completable のいずれかが欠ける。

### pending を completion と誤認してはいけない条件

- `defer-scheduled`、`defer-retry-start`、`defer-retry-end`、`defer-active-context-fallback-start`、`defer-active-context-fallback-end` は completion trace ではありません。
- active CASE context fallback は「対象 workbook を見失った時の fallback refresh 経路」であり、CASE display completed ではありません。
- pending retry attempts exhausted は display failure / observation であり、completion ではありません。
- pending retry が refresh path を再呼び出しても、`case-display-completed` は orchestration の completion 条件を満たすまで emit しません。

### active CASE fallback truth table

`PendingPaneRefreshRetryService` の active CASE fallback は target-lost resiliency fallback です。tracked workbook を見失った場合でも、active context が CASE として解決できるなら refresh attempt を継続するための経路であり、completion fallback、foreground fallback、display session completion、created CASE display completion の代替経路ではありません。

| tracked workbook exists | active context is CASE | attempts remaining | refresh attempted | refresh target | timer continues | completion meaning | trace / outcome meaning |
| --- | --- | --- | --- | --- | --- | --- | --- |
| true | 該当なし | yes | yes | tracked workbook + resolved pane window | refresh success なら stop。refresh failure かつ attempts が残る場合だけ継続。 | active fallback 自体は completion を emit しない。tracked workbook refresh success も completion ではない。 | `defer-retry-start` / `defer-retry-end` は workbook-target retry attempt の観測。`refreshed=true` は refresh attempt result。 |
| false | true | yes | yes | active CASE context | refresh success なら stop。refresh failure かつ attempts が残る場合だけ継続。 | active fallback 自体は completion を emit しない。active CASE fallback success も completion ではない。 | `defer-active-context-fallback-start` / `defer-active-context-fallback-end` は target-lost resiliency fallback の観測。 |
| false | false | yes | no | none | stop | active fallback 自体は completion を emit しない。stop は completion ではない。 | `defer-active-context-fallback-stop` は fallback 不成立の観測。success / recovered / foreground を意味しない。 |
| any | any | no | no | none | stop | active fallback 自体は completion を emit しない。attempts exhausted は completion ではない。 | attempts exhausted による stop は retry lifecycle の観測。completion trace ではない。 |

固定する stop conditions:

- tracked workbook route または active CASE fallback route の refresh success。
- attempts exhausted。
- active context が CASE でない。
- tracked workbook / active fallback のどちらでも refresh attempt できない。

stop の読み替え禁止:

- stop = recovered ではありません。
- stop = foreground success ではありません。
- stop = display session completion ではありません。
- stop = `case-display-completed` ではありません。

trace / outcome の freeze line:

- trace source string / trace payload / trace 名を変更しません。
- `defer-retry-start` / `defer-retry-end` / `defer-active-context-fallback-start` / `defer-active-context-fallback-end` / `defer-active-context-fallback-stop` は observation trace です。
- `refreshed=true` は refresh attempt result であり、単独では completion trace ではありません。
- pending retry success は、existing refresh / outcome / completion chain に戻れたという材料に留まり、`case-display-completed` は display session owner の条件を満たした場合だけ emit できます。

### Phase 4 合格基準

- pending retry `400ms / 3 attempts` と active CASE context fallback が削られていないこと。
- ready-show exhaustion 前に pending retry が始まらないこと。
- pending retry owner を分けても、completion owner が pending retry 側へ移っていないこと。

## 3. WindowActivate Dispatch != Completion

### 固定する意味

`WindowActivate` は downstream refresh の契機です。表示完了そのものではありません。

固定する境界:

- event capture は `ThisAddIn` / `WorkbookEventCoordinator` の境界です。
- request 化、case protection、external workbook detection、case pane suppression、dispatch は `WindowActivatePaneHandlingService` の境界です。
- downstream refresh outcome と created CASE display completion への接続は `TaskPaneRefreshOrchestrationService` の境界です。

`WindowActivateDispatchOutcome.Dispatched` が意味するのは、refresh/display entry へ渡したことだけです。pane visible、visibility recovery completed、foreground guarantee completed、`case-display-completed` のいずれも意味しません。

### trace 上の禁止

- `display-refresh-trigger-dispatched` を display success と読まない。
- `window-activate-display-refresh-trigger-start` / `window-activate-display-refresh-trigger-outcome` を completion trace と読まない。
- WindowActivate 由来の refresh が success しても、created CASE display session の completion 条件を満たすまでは `case-display-completed` とみなさない。
- `WindowActivate` を foreground recovery owner、visibility recovery owner、hidden cleanup owner、white Excel prevention owner にしない。

### Phase 4 合格基準

- WindowActivate observation / dispatch trace が残っていること。
- `Dispatched != completion` の契約が tests / docs / trace naming 上で崩れていないこと。
- WindowActivate の gate 順 `case protection -> external workbook detection -> case pane suppression -> refresh dispatch` が維持されていること。

## 4. Foreground Outcome 条件

### outcome の意味

foreground guarantee は、pane visible 後に foreground obligation が残るかを terminal outcome にする protocol です。

現行 status の freeze line:

| status | freeze line 上の意味 | completion への寄与 |
| --- | --- | --- |
| `RequiredSucceeded` | foreground recovery execution が attempted かつ recovered。 | `IsTerminal == true` かつ `IsDisplayCompletable == true` のため、他条件が揃れば completion に寄与する。 |
| `RequiredDegraded` | foreground recovery execution は attempted したが recovered が false。現行 code では display-completable な degraded terminal。 | `IsTerminal == true` かつ `IsDisplayCompletable == true` のため、他条件が揃れば completion に寄与する。ただし success へ読み替えない。 |
| `NotRequired` | refresh success / pane visible 後でも foreground execution の required 条件が揃わない、または execution 不要。 | `IsTerminal == true` かつ `IsDisplayCompletable == true` のため、他条件が揃れば completion に寄与する。 |
| `SkippedAlreadyVisible` | already-visible path により foreground obligation を satisfied として扱う。 | `IsTerminal == true` かつ `IsDisplayCompletable == true` のため、他条件が揃れば completion に寄与する。 |
| `SkippedNoKnownTarget` | known target がなく foreground guarantee を completion 材料にできない skip。 | `IsDisplayCompletable == false` のため completion に寄与しない。 |
| `RequiredFailed` | vocabulary / 型として存在する non-display-completable failure。 | `IsDisplayCompletable == false` のため completion に寄与しない。現行 required execution false の observed path を勝手にここへ読み替えない。 |
| `Unknown` | 未評価 / execution pending / insufficient facts。 | terminal ではなく completion に寄与しない。 |

### required 判定の固定

foreground execution を required とみなせる条件は、次が揃った場合に限定します。

- refresh が success している。
- pane visible である。
- refresh completed である。
- foreground window がある。
- foreground recovery service が利用可能である。

これらが欠ける場合は、現行条件に従い `NotRequired`、`SkippedNoKnownTarget`、`Unknown` などへ正規化します。欠けた fact を推測で補って foreground execution を広げません。

### foreground fallback の意味

- `ActiveWorkbookFallback` は foreground target kind の fallback であり、completion の fallback ではありません。
- active workbook fallback を使った場合でも、outcome は `RequiredSucceeded` / `RequiredDegraded` などの terminal status へ正規化されてから completion 判定に入ります。
- foreground fallback の存在だけを success と呼びません。

### observation に留まるもの

- foreground trace が emitted されたこと。
- `foreground-recovery-decision` が出たこと。
- `final-foreground-guarantee-started` が出たこと。
- `WindowActivate` が発火したこと。
- `ExcelWindowRecoveryService` が raw facts を返したこと。

これらは outcome の材料または観測であり、単独では completion ではありません。

### Phase 4 合格基準

- foreground decision / outcome / trace owner は `TaskPaneRefreshOrchestrationService` に残っている、または同等の protocol owner に一体で移っていること。
- execution primitive owner (`ExcelWindowRecoveryService`) と decision owner を混ぜていないこと。
- `RequiredDegraded` を success / failure の別意味へ読み替えていないこと。

## R09 Window Resolver / activateWorkbook Route Matrix

### R09 freeze line

R09 は `workbook pane window resolve boundary` です。`ResolveWorkbookPaneWindow(...)` は単なる null guard ではなく、TaskPane display recovery における window availability boundary として扱います。

固定する意味:

- R09 は workbook/window resolve route と `activateWorkbook` flag の route-specific intent を記述する領域です。
- R09 は completion 判定 owner ではありません。
- R09 は foreground decision owner ではありません。
- R09 は retry success owner ではありません。
- `activateWorkbook=true` は activation request であり、foreground guarantee success ではありません。
- `window != null` は display completed ではありません。
- `window == null` は fallback start / retry success / completion failure のいずれにも自動変換しません。

R09 runtime extraction は禁止します。少なくとも Phase 5 で R07 / R13 / R14 / display recovery protocol に触る前に、この route matrix と completion / foreground / retry の距離を維持したまま protocol-preserving shrink として再評価します。

### route matrix

| route | `activateWorkbook` | window unresolved の freeze line | retry / fallback | completion | foreground |
| --- | --- | --- | --- | --- | --- |
| ready-show immediate / wait-ready path | `true` | attempt の window availability が未成立。worker は completion しない。 | `attempt 1 -> 80ms attempt 2 -> pending fallback` だけを維持。 | callback 後の normalized outcome chain を満たす場合だけ。 | activation request は foreground guarantee ではない。 |
| R07 `ScheduleWorkbookTaskPaneRefresh(...)` immediate refresh | `false` | immediate refresh の input window がないだけ。refresh path が fail-closed / context 解決する。 | immediate refresh が success しない場合だけ pending `400ms / 3 attempts`。 | immediate success を completion と読まない。 | なし。 |
| R08 pending retry tick | `true` | retry tick の target workbook window がまだ unavailable。 | attempts を消費して継続、または active CASE context fallback。 | retry success を completion と読まない。 | activation request は foreground success ではない。 |
| `WorkbookOpen` skip / stabilization boundary | 該当なし | `WorkbookOpen + workbook exists + window unresolved` は window-dependent refresh を即時実行しない。 | R07 では pending start へ進まず後続イベントへ委ねる。 | skip は success / completion / fallback start ではない。 | なし。 |
| `WindowActivate` downstream observation | event window を使う。補完時も `false` | WindowActivate dispatch / downstream trace は route observation。 | WindowActivate dispatch 自体は pending fallback ではない。 | `Dispatched != completion`。 | WindowActivate は foreground owner ではない。 |
| `KernelHomeForm.OpenSheet.PostClose` delayed refresh | immediate prepare は `false`、pending tick は `true` | HOME post-close delayed refresh の target window がまだ unavailable。 | R07 immediate 後、必要なら R08 pending retry。 | created CASE reason ではないため `case-display-completed` emit route ではない。 | なし。 |
| foreground guarantee path | 該当なし | foreground decision に必要な window fact が欠ける。 | retry / fallback owner ではない。 | foreground terminal / display-completable は completion 条件の一部。 | R13 / Phase 5 候補。R09 と同一視しない。 |

### WorkbookOpen stabilization boundary

`WorkbookOpen` は window 安定境界ではありません。この freeze line では次を不変とします。

- `WorkbookOpen + workbook exists + window unresolved` の場合は、window-dependent refresh を即時実行しません。
- 後続イベント、`WindowActivate`、または route が許す pending retry route へ委ねます。
- `WorkbookOpen` skip は incidental null guard ではなく stabilization contract です。
- skip を success / completion / fallback start と読みません。
- `WorkbookOpen` 直後の `ActiveWorkbook` / `ActiveWindow` を根拠に window resolve、表示、前面化、pane 対象決定を確定しません。

### WindowActivate / pending retry / foreground との境界

- `WindowActivate` dispatch は completion ではありません。
- `WindowActivateDownstreamObservation` は window availability / route observation の一部であり、`case-display-completed` emit owner ではありません。
- R08 pending retry tick で `activateWorkbook=true` が使われても、目的は window resolve / refresh attempt であり、foreground guarantee success や completion を意味しません。
- foreground guarantee path と window resolve を混同しません。`window resolve`、`activation request`、`foreground outcome`、`display-completable terminal` は別の protocol fact です。

## 5. case-display-completed Emit 条件

### emit してよい条件

`case-display-completed` は、次をすべて満たす場合だけ emit できます。

- reason が created CASE display reason である。現行では `KernelCasePresentationService.ShowCreatedCase.PostRelease` に対応する reason です。
- `TaskPaneRefreshAttemptResult` が存在する。
- `attemptResult.IsRefreshSucceeded == true`。
- `attemptResult.IsPaneVisible == true`。
- `VisibilityRecoveryOutcome` が存在する。
- `VisibilityRecoveryOutcome.IsTerminal == true`。
- `VisibilityRecoveryOutcome.IsDisplayCompletable == true`。
- `attemptResult.IsForegroundGuaranteeTerminal == true`。
- `ForegroundGuaranteeOutcome` が存在する。
- `ForegroundGuaranteeOutcome.IsDisplayCompletable == true`。
- created CASE display session が解決できる。
- その session がまだ completed ではない。

### emit してはいけない条件

- `WorkbookTaskPaneReadyShowAttemptWorker`、`TaskPaneRefreshCoordinator`、`TaskPaneHostFlowService`、`TaskPaneManager` から直接 emit しない。
- `taskpane-already-visible` だけを根拠に emit しない。
- `taskpane-refresh-completed` だけを根拠に emit しない。
- `WindowActivateDispatchOutcome.Dispatched` だけを根拠に emit しない。
- pending retry scheduled / exhausted を根拠に emit しない。
- visibility / foreground outcome が `Unknown`、non-terminal、non-display-completable の場合は emit しない。
- created CASE reason ではない refresh route から emit しない。

### one-time emit の意味

- created CASE display session は workbook full name と session id で追跡されます。
- emit は session ごとに 1 回だけです。
- emit 成功時に session は completed とされ、active sessions から外れます。
- duplicate emit を防ぐ状態管理は completion owner の一部であり、単なる state bag helper ではありません。

### ready-show / already-visible / fallback path の収束条件

- ready-show refresh path、already-visible path、pending retry fallback path は、どれも raw facts を返すだけでは completion ではありません。
- どの path から戻っても、completion は `TaskPaneRefreshOrchestrationService.TryCompleteCreatedCaseDisplaySession(...)` 相当の条件に収束します。
- already-visible path では refresh render が走らない場合があります。この場合も `TaskPaneRefreshAttemptResult.VisibleAlreadySatisfied()` のような success 相当 facts が visibility / foreground outcome に正規化され、display-completable 条件を満たした場合だけ completion できます。
- fallback path では retry success が completion ではなく、retry 後の normalized outcome が completion 条件を満たすことが必要です。

### Phase 4 合格基準

- `case-display-completed` の emit owner が分散していないこと。
- one-time emit state が helper 化により意味を失っていないこと。
- lower-level owner が pane visible / refresh completed を display completed と名乗っていないこと。

## 6. Display Session Boundary

### created CASE display session の開始条件

created CASE display session は、ready-show request を `TaskPaneRefreshOrchestrationService` が受理した時点で開始します。

開始できる条件:

- reason が created CASE display reason である。
- workbook が null ではない。
- workbook full name が取得できる。

開始時に保持する情報:

- `sessionId`
- `workbookFullName`
- `reason`
- completed state

開始時に emit する trace:

- `created-case-display-session-started`
- `display-handoff-completed`

### 終了条件

session の終了は `case-display-completed` の one-time emit です。

終了できる条件:

- `case-display-completed` emit 条件をすべて満たす。
- session が未完了である。

終了後:

- session は completed になる。
- active session tracking から外れる。
- 同じ session id の duplicate completion は emit しない。

### route をまたぐときに保持すべき情報

ready-show、already-visible、refresh、pending retry、WindowActivate downstream trace をまたぐ場合も、次の相関を失ってはいけません。

- created CASE display reason。
- workbook full name。
- session id。
- completion source。
- attempt number。
- pane visible facts。
- visibility recovery outcome。
- refresh source outcome。
- rebuild fallback outcome。
- foreground guarantee outcome。
- WindowActivate trigger facts は completion ではなく observation として保持する。

### session を helper 化してはいけない理由

display session は単なる dictionary / state bag ではありません。

- `display-handoff-completed` と `case-display-completed` の対応関係を守ります。
- ready-show acceptance と final completion を同じ protocol に閉じます。
- already-visible path と refresh path を同じ completion 条件へ収束させます。
- one-time emit を保証します。
- foreground terminal / visibility terminal が揃うまで completion しない fail-closed 条件を持ちます。

したがって Phase 4 で R04 / R14 を動かす場合、session state だけを外へ出すのではなく、display protocol session boundary として一体で扱います。R04 / R14 は Phase 4 後半まで orchestration に残すのが安全です。

## 7. Protection Gate / Fail-Closed 条件

### gate を通過しない場合の扱い

protection / suppression / precondition gate を通過しない場合、refresh success や display completion に丸めません。

固定する gate:

- `WorkbookOpen` 直後の window-dependent refresh skip。
- case foreground protection 中の `WorkbookActivate` ignore。
- case foreground protection 中の `WindowActivate` ignore。
- case foreground protection 中の `TaskPaneRefresh` ignore。
- case pane activation suppression 中の refresh defer / skip。
- external workbook detection による display request 抑止。
- `_workbookSessionService` / `_taskPaneManager` など required dependency missing。
- taskPane refresh suppression count > 0。
- context rejected。
- unknown role / missing window key による host-flow hide / skip。

### fail-closed と retry / fallback の関係

- fail-closed は retry 開始条件ではありません。
- fail-closed result を success に丸めて pending retry を終わらせません。
- protection 中に無理な refresh を通して retry / fallback を消化しません。
- `WorkbookOpen` window-dependent skip は後続 `WorkbookActivate` / `WindowActivate` へ委ねる境界であり、推測で window を補って retry しません。
- pending retry は fallback scheduling owner であり、fail-closed 条件の意味を変更しません。

### UI policy との関係

この freeze line は `docs/ui-policy.md` の次の原則に従います。

- WorkbookOpen 直接依存の表示制御を行わない。
- 表示は専用サービス経由で行う。
- ScreenUpdating は必ず復元する。
- Window 状態は復旧処理を前提とする。
- TaskPane は遅延表示を前提とする。
- hidden session を表示制御の一般手段として使わない。
- `Application.DoEvents()`、sleep、timing hack を追加しない。

### Phase 4 合格基準

- protection / suppression / fail-closed gate を通過しない route が completion として扱われていないこと。
- protection gate の active window 基準を、実機観測なしに narrower / broader へ変更していないこと。
- fail-closed 条件を retry 値や fallback owner の分離に混ぜて変更していないこと。

## 8. Trace Contract

### trace 名を変えてはいけない理由

TaskPane 表示回復の trace 名は、実機表示不安定、白 Excel、foreground degradation、WindowActivate downstream、ready-show fallback を切り分ける観測契約です。

trace 名を変えると、次が壊れます。

- Phase 0/1/2 docs との対応。
- 実機観測 checklist との対応。
- `WindowActivate dispatch != completion` の誤認防止。
- ready-show / pending retry / foreground / completion の route 切り分け。
- 将来 refactor の挙動不変確認。

### 意味を変えてはいけない trace

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

### downstream trace と completion trace の境界

- downstream trace は、trigger / dispatch / refresh route の観測です。
- completion trace は、created CASE display session が terminal display-completable 条件を満たしたことの観測です。
- `window-activate-display-refresh-trigger-outcome` は downstream trace であり、completion trace ではありません。
- `case-display-completed` だけが created CASE display completion trace です。

### Phase 4 合格基準

- trace 名を rename していないこと。
- trace の source owner を動かす場合も、意味と発火条件を変えていないこと。
- downstream trace と completion trace を同じ event として扱っていないこと。

## Phase 4 Safe-First Ownership 分離への渡し

Phase 4 safe-first の安全領域:

1. R02 refresh precondition / fail-closed policy boundary は Phase 4 最初の safe unit として完了済み。
2. R16 timer lifecycle boundary の owner 明確化は Phase 4 R16 safe unit で完了済み。`TaskPaneRetryTimerLifecycle` が ready-show retry timer と pending retry timer の create / register / stop / unregister / dispose を持ち、retry 順序・pending 条件・completion 条件は変更しない。
3. R06 ready-show retry scheduler の owner 明確化は Phase 4 R06 safe unit で完了済み。`TaskPaneReadyShowRetryScheduler` が 80ms retry scheduling と scheduled / firing trace emission を持ち、`TaskPaneDisplayRetryCoordinator` は attempt sequencing、`WorkbookTaskPaneReadyShowAttemptWorker` は attempt 本体、`TaskPaneRetryTimerLifecycle` は timer lifecycle、`TaskPaneRefreshOrchestrationService` は ready-show acceptance / callback / fallback handoff / completion への収束を維持する。`attempt 1 -> 80ms retry attempt 2 -> pending retry fallback`、pending retry `400ms / 3 attempts`、callback 意味、completion 条件、display session boundary、trace 名と意味は変更しない。
4. R10/R11/R12 normalized outcome mapping の decision object 化検討。
5. R15 WindowActivate downstream observation boundary の owner 明確化は Phase 4 R15 safe unit で完了済み。`WindowActivateDownstreamObservation` が `window-activate-display-refresh-trigger-start` / `window-activate-display-refresh-trigger-outcome` と WindowActivate display request trace fields を持ち、`WindowActivatePaneHandlingService` は dispatch gate、`TaskPaneRefreshOrchestrationService` は refresh path ordering / completion 接続 / display session boundary を維持する。`WindowActivate dispatch != completion`、`case-display-completed` one-time emit、foreground outcome semantics、trace 名と意味、callback 意味、retry sequencing は変更しない。

Phase 4 でもまだ触ってはいけない領域:

- R04 / R14 display protocol session boundary。
- R05 ready-show callback と R10/R13/R14 completion 収束。
- R06/R07/R08 retry sequence の順序変更。
- R09 `activateWorkbook` route policy。
- R13 foreground guarantee decision と completion 接続。
- `case-display-completed` emit owner。

## 今回行わないこと

- コード変更。
- service 新設。
- helper 抽出。
- class rename。
- namespace 移動。
- method move。
- retry 値 / 順序変更。
- trace 名変更。
- route contract 変更。
- fail-closed 条件変更。
- foreground / visibility recovery 条件変更。
- completion 条件変更。
- UI policy 変更。
- COM restore 順序変更。
- build / test / `DeployDebugAddIn` 実行。
