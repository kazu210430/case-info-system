# WorkbookClose / Reopen Protocol Current Mapping

## 位置づけ

この文書は、hidden Excel / isolated app / white Excel lifecycle redesign の E-0 として、WorkbookClose / reopen protocol の current mapping を docs-only で固定するためのものです。

- 基準 main / origin/main: `dc9a71b1b726a5eec40081c87b7756b9e4b1e01d`
- 対象: WorkbookClose / managed close / unmanaged close / user close / system close / post-close follow-up / reopen / close 後参照境界
- 非対象: 条件変更、実装変更、test 変更、build / test / Debug Add-in 配備

参照した正本:

- `AGENTS.md`
- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/case-workbook-lifecycle-current-state.md`
- `docs/hidden-excel-isolated-app-white-excel-lifecycle-current-state.md`
- `docs/hidden-excel-isolated-app-white-excel-lifecycle-target-state.md`
- `docs/hidden-excel-lifecycle-outcome-vocabulary.md`
- `docs/case-display-recovery-protocol-current-state.md`
- `docs/case-display-recovery-protocol-target-state.md`

この文書は current-state の棚卸しだけを扱います。WorkbookClose 条件、reopen 条件、post-close follow-up 条件、foreground / visibility 条件、hidden cleanup / retained cleanup 条件、WindowActivate dispatch 条件は変更しません。

## current mapping summary

- WorkbookClose event は `ThisAddIn.Application_WorkbookBeforeClose(...)` で受け、`WorkbookLifecycleCoordinator.OnWorkbookBeforeClose(...)` が CASE / Kernel / Accounting lifecycle へ順に渡します。
- CASE close 判定の中心 owner は `CaseWorkbookLifecycleService` です。`CaseWorkbookBeforeClosePolicy` は out-of-scope / managed close / dirty / clean の current outcome を決めます。
- dirty user close は一度 cancel され、close 前に workbook key と folder path を確定し、prompt / folder offer / managed close / post-close follow-up へ進みます。
- managed close は `ManagedCloseState` scope 内で prompt を抑止し、対象 workbook を key から再取得して、`WorkbookCloseInteropHelper.CloseWithoutSave(workbook)` で optional 引数を明示して close します。
- clean / unmanaged close は close を cancel せず、close 前に workbook key と folder path を確定して post-close follow-up を予約します。
- post-close follow-up は closed workbook object を使わず、workbook key と add-in 側が保持する current `Application` から still-open / visible workbook を判定します。
- white Excel prevention は `PostCloseFollowUpScheduler` の責務です。visible workbook が無い場合だけ `Application.Quit()` を試行し、visible workbook がある場合は quit しません。
- interactive created CASE の reopen は、hidden create session close 後に `KernelCasePresentationService` が `CaseWorkbookOpenStrategy.OpenHiddenForCaseDisplay(...)` を呼び、shared/current app 上で新しい workbook object を取得する流れです。
- reopen 後の visibility / recovery / foreground / TaskPane completion は display / refresh protocol 側の責務です。WorkbookClose owner や post-close follow-up owner ではありません。
- `WindowActivate` は display / refresh dispatch trigger です。close 条件、reopen 判断、hidden cleanup、retained cleanup、white Excel prevention、foreground guarantee の owner ではありません。
- white Excel prevention / recovery の G-0 current-state と target boundary は `docs/white-excel-prevention-boundary-current-state.md` を参照します。

## owner / trigger / primitive / trace owner

### WorkbookClose

| protocol unit | owner | trigger | primitive | trace owner |
| --- | --- | --- | --- | --- |
| event boundary | `ThisAddIn` | Excel `WorkbookBeforeClose` | `EventBoundaryGuard.ExecuteCancelable(...)` | `WorkbookEventLogger` / `WorkbookLifecycleCoordinator` |
| close orchestration | `WorkbookLifecycleCoordinator` | `ThisAddIn.HandleBeforeClose(...)` | CASE -> Kernel -> Accounting -> TaskPane cleanup の順次委譲 | `WorkbookEventLogger` |
| CASE close policy | `CaseWorkbookLifecycleService` / `CaseWorkbookBeforeClosePolicy` | coordinator からの workbook close | role 判定、managed state、dirty state | `CaseWorkbookLifecycleService` |
| managed close state | `ManagedCloseState` | dirty close 後の scheduled managed close | workbook key scope | service-side trace only |
| managed close primitive | `CaseWorkbookLifecycleService` / `WorkbookCloseInteropHelper` | dispatcher 経由の scheduled close | `CloseWithoutSave(workbook)` with `false, Type.Missing, Type.Missing` | `CaseWorkbookLifecycleService` |
| post-close follow-up | `PostCloseFollowUpScheduler` | close 前予約後の UI queue / retry timer | workbook key、current `Application.Workbooks` enumeration、visible window scan、`Application.Quit()` | `WhiteExcelPrevention*` trace |
| TaskPane host cleanup | `WorkbookLifecycleCoordinator` / `TaskPaneManager` / `TaskPaneHostLifecycleService` | close が cancel されなかった `WorkbookBeforeClose` | close 前 workbook full name による host removal | TaskPane lifecycle trace |
| hidden create session close | `CaseWorkbookOpenStrategy` | created CASE save / cleanup | hidden workbook close、hidden Application quit / release | hidden workbook cleanup / isolated app outcome trace |
| retained hidden app cleanup | `CaseWorkbookOpenStrategy` | return-to-idle、poison、shutdown、idle timeout | cached workbook close、cached Application retain / quit / release | retained instance cleanup outcome trace |

### reopen

| protocol unit | owner | trigger | primitive | trace owner |
| --- | --- | --- | --- | --- |
| interactive created CASE reopen | `KernelCasePresentationService` | `OpenCreatedCase(...)` after successful create result | `CaseWorkbookOpenStrategy.OpenHiddenForCaseDisplay(...)` | case workbook open / presentation trace |
| shared/current app reopen | `CaseWorkbookOpenStrategy` | presentation owner request | `_application.Workbooks.Open(path)`、temporary hidden windows、previous window restore、shared app state restore | open strategy trace |
| non-hidden open route | `CaseWorkbookOpenStrategy` | presentation owner request for non-default route | `OpenVisibleWorkbook(...)` current behavior | open strategy trace |
| display handoff | `KernelCasePresentationService` | fresh workbook object after reopen | `WorkbookWindowVisibilityService`、`ExcelWindowRecoveryService`、ready-show request | presentation / recovery trace |
| foreground guarantee decision | `TaskPaneRefreshOrchestrationService` | ready-show / refresh protocol | foreground guarantee outcome vocabulary | refresh orchestration trace |
| activation refresh dispatch | `WindowActivatePaneHandlingService` | Excel `WindowActivate` captured facts | `TaskPaneDisplayRequest.ForWindowActivate(...)` | `WindowActivateDispatchOutcome` trace |

## managed / unmanaged / user / system / post-close 境界

| close kind | current-state boundary |
| --- | --- |
| managed close | `ManagedCloseState` が workbook key scope を持つ close。CASE dirty prompt 後の scheduled close では prompt を抑止し、save intent と post-close follow-up を close 前に確定します。 |
| unmanaged close | `ManagedCloseState` scope 外の close。CASE clean close は cancel せず、post-close follow-up を予約して Excel の close に進めます。 |
| user close | user 操作由来かどうかは Excel event だけからは厳密に正規化されていません。CASE dirty close は prompt によって user choice を取得します。 |
| system close | Excel shutdown / add-in shutdown 由来の close を専用 vocabulary として正規化する current owner は未定義です。`WorkbookBeforeClose` が発火した範囲では同じ coordinator に入ります。`ThisAddIn_Shutdown` は retained hidden app-cache cleanup の呼び出し境界です。 |
| post-close follow-up | close 前に確定した workbook key を使う deferred queue です。closed workbook object の状態確認ではなく、current `Application.Workbooks` 列挙による still-open check と visible workbook scan を行います。 |

## owner / trigger / guarantee / recovery / trace / UX

| area | owner | trigger | guarantee | recovery | trace | UX |
| --- | --- | --- | --- | --- | --- | --- |
| CASE close policy | `CaseWorkbookLifecycleService` | `WorkbookBeforeClose` | out-of-scope / managed / dirty / clean の分岐 | prompt cancel / prompt exception は close cancel 状態を戻す | lifecycle logs | dirty prompt、folder offer |
| managed close execution | `CaseWorkbookLifecycleService` | dispatcher scheduled action | close 前情報を使って close を試行する。close 成功後の workbook 再参照はしない | close failure は log と message | close error / lifecycle logs | failure dialog title は close 失敗時のみ |
| post-close white Excel prevention | `PostCloseFollowUpScheduler` | queued follow-up / retry | workbook が閉じており visible workbook が無い場合だけ quit を試行 | Excel busy retry、quit failure log、failure 時 DisplayAlerts restore | `WhiteExcelPreventionQueued` / `NotRequired` / `Completed` / `Failed` | failure 時の追加 UX は current-state では未定義 |
| reopen for display | `KernelCasePresentationService` / `CaseWorkbookOpenStrategy` | created CASE success | fresh workbook object を shared/current app から取得し、shared app state を復元する | open failure は opened workbook cleanup と previous window restore | open / display trace | wait UI stage、ready-show |
| visibility / foreground | `WorkbookWindowVisibilityService` / `ExcelWindowRecoveryService` / `TaskPaneRefreshOrchestrationService` | display handoff / ready-show / refresh | close protocol ではなく display protocol の outcome として記録 | recovery primitive と retry 側 | display / refresh trace | visible pane / foreground |
| WindowActivate dispatch | `WindowActivatePaneHandlingService` | Excel `WindowActivate` | display refresh request を dispatch / defer / ignore する | recovery owner ではない | `WindowActivateDispatchOutcome` | direct UX owner ではない |

## close 前に capture すべき information

current-state 上、close 後に workbook COM object を安全に読める前提を置かないため、以下は close 前に確定する境界として扱います。

- workbook identity: workbook key、full path / name、known case path alias
- folder UX facts: created case folder offer pending、containing folder path
- close policy facts: CASE / Base / Kernel / Accounting 対象判定、managed close scope、dirty session state
- user decision facts: dirty prompt result、save intent、folder offer result
- scheduling facts: post-close follow-up に渡す workbook key、retry queue の識別情報
- TaskPane cleanup facts: target workbook の full name / host key
- hidden session facts: hidden workbook path、hidden Application lifetime owner、cache slot ownership、poison / return-to-idle decision
- reopen handoff facts: created CASE path、operation kind、wait session、transient suppression path

post-close follow-up の visible workbook scan は close 後に current `Application` から取得する事実であり、閉じた workbook から capture する情報ではありません。

## close 後に再参照してはいけない information

WorkbookClose 後の current mapping では、少なくとも次の再参照を protocol 上の危険単位として扱います。

- closed workbook COM object の `FullName` / `Name` / `Saved` / `Application` / `Windows` / `Sheets` / `CustomDocumentProperties`
- closed workbook から辿った `Application`、`ActiveWindow`、window visibility、pane host
- close 済み window を前提にした `TaskPane` host lookup / refresh / activation
- close 済み workbook の DocProperty / role / `SYSTEM_ROOT` / snapshot / cache state
- close 済み workbook を使った reopen 判断

current implementation の post-close follow-up は workbook key を使って current `Application.Workbooks` を列挙します。これは closed workbook object の再参照ではありません。managed close failure path では close が成功した前提に入っていないため、failure dialog の title 解決に workbook を参照し得ますが、close 成功後の再参照とは区別します。

## reopen 時に再取得すべき information

reopen は close 前 workbook object の継続利用ではなく、新しい workbook object / window facts の取得境界です。reopen 時は次を再取得対象として扱います。

- fresh workbook COM object
- workbook windows、active window、visible state
- workbook role / known case path / base-case 判定
- DocProperty / snapshot / `SYSTEM_ROOT` / generated file facts
- TaskPane host / pane visibility / display request context
- foreground guarantee に必要な workbook / window / application facts
- transient suppression / protection / ready-show に必要な current facts
- previous window restore の成否と shared app state restore outcome

## post-close follow-up と white Excel prevention

G-0 の詳細な current meaning は `docs/white-excel-prevention-boundary-current-state.md` を参照します。この節は WorkbookClose / reopen protocol から見た接続点を扱います。

- `PostCloseFollowUpScheduler.Schedule(workbookKey, folderPath)` は close 前に予約され、UI queue で遅延実行されます。
- follow-up はまず workbook key で still-open check を行います。対象 workbook がまだ open の場合、quit しません。
- 対象 workbook が closed と判定された後、open workbooks の visible window を scan します。
- visible workbook がある場合は `WhiteExcelPreventionNotRequired` として quit しません。
- visible workbook が無い場合だけ `DisplayAlerts=false` にして `Application.Quit()` を試行し、成功時は `WhiteExcelPreventionCompleted` を記録します。
- quit 成功時に DisplayAlerts restore は行いません。quit failure 時は snapshot がある場合に DisplayAlerts を restore します。
- Excel busy は retry queue に戻します。retry 上限後または他 exception は failure trace です。
- `folderPath` は current request payload に含まれますが、scheduler 内の white Excel prevention 判定 primitive では使われていません。
- user が WorkbookClose 直後に同じ CASE を即 reopen した場合でも、current-state は follow-up cancel / reopen gating ではなく、queued key と current `Application.Workbooks` 列挙による fail-closed 判定として動きます。正式な cancel / gating protocol は current-state 上の未定義ポイントとして残します。

## E-2 post-close follow-up / immediate reopen current behavior

E-2 では runtime 条件を変えず、post-close follow-up と immediate reopen が近接した場合の current behavior を次のように読むことにします。

- close 前に `CaseWorkbookLifecycleService` が確定するのは queued key です。`workbook-close-immutable-facts-captured` と `workbook-close-follow-up-facts-captured` は close 前 workbook object から既存フローで得ていた facts を記録します。
- reopen は別 owner です。interactive created CASE の reopen では `KernelCasePresentationService` / `CaseWorkbookOpenStrategy` が shared/current app 上で fresh workbook object を取得し、以後の workbook / window / DocProperty / TaskPane facts は reopen 後に再取得する facts です。
- `PostCloseFollowUpScheduler` の still-open 判定は、queued key と current `Application.Workbooks` 列挙から得た fresh open workbook key を比較する判定です。closed workbook object の再参照ではありません。
- immediate reopen 後の fresh workbook が queued key と一致する場合、current behavior では `targetWorkbookStillOpen=True` / `decision=skip-still-open` として quit せず、`WhiteExcelPreventionNotRequired` の `outcomeReason=targetWorkbookStillOpen` を記録します。これは follow-up cancel ではなく、dequeue 時の still-open skip です。
- immediate reopen 後の fresh workbook が queued key と一致しない場合でも、visible workbook scan で visible workbook が見つかれば `WhiteExcelPreventionNotRequired` の `outcomeReason=visibleWorkbookExists` として quit しません。この場合、queued key の target は closed と読まれ、fresh visible workbook の存在だけが quit 抑止理由です。
- still-open が false で、visible workbook も無い場合だけ `Application.Quit()` を試行し、成功時に `WhiteExcelPreventionCompleted` を記録します。quit 成功後に終了中 `Application` を restore しない境界は変えません。
- `WhiteExcelPreventionFailed` は no visible workbook quit が試行され、例外等で完了しなかった場合の failure trace です。foreground recovery、visibility recovery、WindowActivate dispatch で補完しません。

### E-2 classification

| classification | current decision |
| --- | --- |
| E-2a docs 追記のみ | 採用。queued key / fresh reopen facts / still-open / WhiteExcelPrevention outcome の意味を docs で固定します。 |
| E-2b tests 追加のみ | 今回は未採用。既存 `PostCloseFollowUpSchedulerTests` が visible workbook skip、quit success/failure、still-open skip diagnostic を持つため、まず docs で current behavior を固定します。 |
| E-2c diagnostic trace / outcome vocabulary の軽微整合 | 今回は docs vocabulary の意味補足に限定します。runtime trace field や emitted condition は変更しません。 |
| E-2d private helper 化が必要なもの | 未採用。still-open / visible scan の実装分解は行いません。 |
| E-2e follow-up cancel / reopen gating / visible 判定 / Quit 条件 | 禁止。今回の整理では実装しません。 |

E-2 の結果、競合時の current-state は「queued key と fresh open workbook facts を比較するが、queued key 自体は rewrite しない」と表現します。fresh reopen facts は display / refresh protocol 側で再取得され、post-close follow-up request payload の captured facts と混ぜません。

## foreground / visibility との境界

- WorkbookClose protocol は foreground guarantee owner ではありません。
- post-close follow-up は white Excel prevention owner であり、foreground / visibility restore owner ではありません。
- `OpenHiddenForCaseDisplay(...)` は shared/current app の reopen を一時 hidden にして previous window restore までを扱います。
- workbook window を表示可能状態へ寄せる lightweight ensure は `WorkbookWindowVisibilityService` です。
- app/window recovery primitive は `ExcelWindowRecoveryService` です。
- foreground guarantee decision / outcome / trace は display / refresh protocol、特に `TaskPaneRefreshOrchestrationService` 側で扱います。
- `WorkbookOpen` は window 安定境界ではありません。window-dependent refresh は `WorkbookActivate` / `WindowActivate` 以降、または ready-show / retry 側で扱います。

## hidden Excel cleanup / retained cleanup との境界

- hidden create session の close / quit / release mechanics は `CaseWorkbookOpenStrategy` の責務です。
- hidden create session の operation owner は `KernelCaseCreationService` ですが、hidden workbook open / close mechanics と isolated Application cleanup owner は `CaseWorkbookOpenStrategy` です。
- retained hidden app-cache の return-to-idle / poison / shutdown / idle timeout cleanup owner も `CaseWorkbookOpenStrategy` です。
- `ThisAddIn_Shutdown` は retained hidden app-cache の shutdown cleanup を呼ぶ境界です。
- white Excel prevention の `PostCloseFollowUpScheduler` は shared/current app の visible workbook absence に対する quit 判定であり、hidden create session cleanup や retained cached Application cleanup の代替 owner ではありません。
- retained cached Application を user-visible workbook の white Excel prevention と混同しません。

## WindowActivate が関与してはいけない範囲

current-state 上、`WindowActivate` は次の owner ではありません。

- WorkbookClose 条件判定
- managed close / unmanaged close / user close / system close の分類
- close 後の workbook still-open 判定
- post-close follow-up queue / retry / quit 判定
- reopen 条件判定
- hidden create session cleanup
- retained hidden app-cache cleanup
- white Excel prevention
- foreground guarantee outcome の最終 owner
- close 後 COM object 再参照リスクの補正

`WindowActivatePaneHandlingService` は captured facts を display refresh request へ変換し、Observed / Ignored / Deferred / Dispatched / Failed の dispatch outcome を記録する boundary です。

## fail closed 境界

- close 対象 workbook が out-of-scope の場合、CASE close policy は intervention しません。
- dirty prompt cancel は close を cancel し、managed close / post-close follow-up を進めません。
- prompt exception は close cancel を解除して Excel の通常 close を妨げない方向に戻します。
- post-close follow-up で target workbook が still open と判定された場合、quit しません。
- visible workbook が 1 つでもある場合、white Excel prevention は quit しません。
- Excel busy は retry し、retry exhausted / non-busy exception は failure trace に落とします。
- reopen failure は opened workbook cleanup と previous window restore を試み、表示 protocol へ成功扱いで進めません。
- owner が不明な Application / workbook / window / context からは cleanup / foreground / reopen 成功を推論しません。

## current-state 上の未定義ポイント

- user close と system close を event facts だけで区別する正規化 vocabulary は未定義です。
- WorkbookClose 直後に同一 CASE が即 reopen された場合、current behavior は E-2 の still-open / visible workbook scan として整理しました。ただし follow-up cancel、reopen gating、user-facing UX を含む正式な競合 protocol は未定義です。
- `PostCloseFollowUpScheduler` の quit failure 後に user-facing UX を出すかは未定義です。
- scheduler request の `folderPath` は payload に残っていますが、white Excel prevention primitive としての意味は未定義です。
- temporary read-only workbook close、master/kernel helper close、hidden session close を CASE managed close protocol に含めるかは current-state では分けて扱っています。
- close failure path で workbook object をどこまで error reporting に使ってよいかの細かい境界は未定義です。
- foreground guarantee outcome と reopen outcome の相互参照粒度は target-state 側で分離されていますが、current code の trace vocabulary は完全には統一されていません。

## tests / trace current mapping

この E-0 では tests を実行しません。読取ベースで current mapping に関係する test coverage は次の通りです。

- `WorkbookCloseInteropHelperTests`: optional close arguments の固定
- `ManagedCloseStateTests`: nested managed close scope と workbook key 判定
- `CaseWorkbookLifecycleServicePolicyTests`: before-close policy outcome
- `CaseWorkbookLifecycleServiceThinOrchestrationTests`: dirty prompt、managed close scheduling、clean post-close scheduling、folder offer
- `PostCloseFollowUpSchedulerTests`: visible workbook check、quit success/failure、DisplayAlerts restore、busy retry
- `CaseWorkbookOpenStrategyTests`: hidden/session cleanup、retained cleanup、shared app state restore、open failure cleanup
- `TaskPaneHostLifecycleServiceTests`: close 前 workbook key による host removal
- `WindowActivatePaneHandlingServiceTests`: WindowActivate dispatch outcome と non-owner 境界

## E-1 diagnostic trace vocabulary

E-1 では runtime 条件を変えず、close 前 facts と post-close follow-up decision の観測点を追加する。

- `workbook-close-immutable-facts-captured`
  - trace owner: `CaseWorkbookLifecycleService`
  - close 前に既存取得済みの `workbookKey`、`isBaseOrCaseWorkbook`、`isManagedClose`、`isSessionDirty`、`beforeCloseAction` を記録する。
  - `SYSTEM_ROOT`、DocProperty、window、sheet などの追加 COM 参照は行わない。
- `workbook-close-follow-up-facts-captured`
  - trace owner: `CaseWorkbookLifecycleService`
  - close 前に既存フローで解決していた `folderPath` の有無と `beforeCloseAction` を、post-close / managed close へ渡る immutable facts として記録する。
- `post-close-follow-up-request-dequeued`
  - trace owner: `PostCloseFollowUpScheduler`
  - close 前に queue へ積まれた `workbookKey`、`folderPath` 有無、attempts、queue count を request payload として記録する。
- `post-close-follow-up-decision`
  - trace owner: `PostCloseFollowUpScheduler`
  - close 後は closed workbook object ではなく、captured `workbookKey` と current `Application.Workbooks` 列挙による still-open decision を記録する。

これらは diagnostic trace であり、WorkbookClose 条件、reopen 条件、post-close quit 条件、visible workbook 判定条件、foreground / visibility 条件、WindowActivate dispatch 条件を変更しない。

## E-1 で触るべき安全単位

E-1 で触る場合も、条件変更ではなく観測可能性と境界固定を最小単位にします。

- close 前 immutable facts capture の明示化
- close / reopen outcome trace の vocabulary 寄せ
- post-close follow-up queue の diagnostic trace 整理
- closed workbook object を close 後に読まないことを確認する test / trace 単位
- reopen 後に再取得した facts と close 前 facts の区別を trace する単位
- WindowActivate が cleanup / white Excel / foreground guarantee owner ではないことを trace vocabulary で維持する単位

## E-1 で触ってはいけない危険単位

- WorkbookClose 条件
- reopen 条件
- dirty prompt / folder offer / managed close scheduling 条件
- post-close follow-up の still-open / visible workbook / quit / retry 条件
- foreground / visibility restore 条件
- hidden create session cleanup 条件
- retained hidden app-cache cleanup 条件
- WindowActivate dispatch 条件
- `WorkbookCloseInteropHelper` optional arguments
- close 後 workbook COM object / workbook-derived Application / closed window を読む実装
- service / helper の大規模抽出や責務移動
- timing workaround による順序補正

## 今回行わないこと

- コード変更
- tests 変更
- build / test 実行
- `DeployDebugAddIn` 実行
- WorkbookClose / reopen / post-close / visibility / foreground / hidden cleanup / retained cleanup / WindowActivate dispatch の条件変更
