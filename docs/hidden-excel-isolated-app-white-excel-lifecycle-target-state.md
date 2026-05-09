# Hidden Excel / Isolated App / White Excel Lifecycle Target State

## 位置づけ

この文書は、hidden Excel / isolated app / retained hidden app-cache / white Excel lifecycle の target-state を、docs-only で固定する設計記録です。

- current-state 正本化済み基準:
  - `2026-05-09` 作業開始時点で `main` / `origin/main` / `HEAD` が一致した `e2c7afa49b5e9a736396b933ba5500bac5311ec7`
- 参照した正本:
  - `AGENTS.md`
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
  - `docs/case-display-recovery-protocol-current-state.md`
  - `docs/case-display-recovery-protocol-target-state.md`
  - `docs/case-workbook-lifecycle-current-state.md`
  - `docs/hidden-excel-isolated-app-white-excel-lifecycle-current-state.md`

この文書は実装を変更しません。コード変更、helper / service 抽出、Excel visibility 制御変更、hidden Excel cleanup 条件変更、WorkbookClose / reopen 条件変更、foreground / visibility / rebuild / refresh source 条件変更、build / test / `DeployDebugAddIn` 実行は行いません。

## Target-State Summary

target-state では、hidden Excel / isolated app / retained hidden app-cache / white Excel lifecycle を、表示回復の一部ではなく、lifecycle owner と protocol owner の集合として扱います。

- shared/current app は原則 user-owned です。業務処理側は quit owner ではありません。
- isolated app は生成した owner が close / quit / COM release まで閉じます。
- retained hidden app-cache は `CaseWorkbookOpenStrategy` の例外境界です。session close と cached `Application` cleanup を混同しません。
- hidden workbook window の visibility restore は保存状態正規化または display handoff の準備です。retained instance cleanup や foreground guarantee の代替ではありません。
- white Excel 防止は close / quit 側の protocol です。`WindowActivate`、visibility recovery、foreground guarantee の代替ではありません。
- `WindowActivate` は display / refresh trigger に限定します。cleanup / recovery / foreground guarantee owner ではありません。
- `WorkbookClose` 後は、閉じた workbook を再参照しない fail-closed 境界を維持します。

## Lifecycle Owner Separation

| lifecycle unit | target-state owner | owner が持つこと | owner が持たないこと |
| --- | --- | --- | --- |
| Excel instance lifetime owner | shared/current app は user / Excel host。isolated app は生成 service。retained cached app は `CaseWorkbookOpenStrategy`。 | 自分が生成した `Application` の close / quit / release 境界を閉じる。shared/current app は no visible workbook quit の例外を除き終了しない。 | 他 owner の `Application` を quit すること。`WindowActivate` に lifetime を委譲すること。 |
| Excel visibility owner | display/recovery 側は `WorkbookWindowVisibilityService` / `ExcelWindowRecoveryService`。Kernel HOME は `KernelWorkbookDisplayService`。保存状態正規化は session owner。 | workbook window visible ensure、application/window recovery、保存前 visibility normalization をそれぞれの用途内で行う。 | visibility restore を retained cleanup や foreground guarantee の代替にすること。 |
| hidden Excel cleanup owner | hidden session を開始した owner。CASE create mechanics は `CaseWorkbookOpenStrategy`、reflection は `KernelUserDataReflectionService`。 | owned workbook close、owned isolated app quit、COM release、cache return / poison 判定。 | shared/current app の user-owned workbook / app を閉じること。 |
| retained instance cleanup owner | `CaseWorkbookOpenStrategy` hidden app-cache。 | idle return、timeout、poison、feature flag disabled、shutdown cleanup。 | user-owned shared app を閉じること。session owner 以外が cached app を破棄すること。 |
| WorkbookClose / reopen protocol owner | close は `CaseWorkbookLifecycleService` / `ManagedCloseState` / `PostCloseFollowUpScheduler`。reopen は `KernelCasePresentationService` / `CaseWorkbookOpenStrategy`。 | close 前 facts 取得、managed close、post-close follow-up、interactive CASE reopen。 | close 後 workbook 再参照。reopen 時に古い workbook / window / context を使い回すこと。 |
| foreground guarantee owner | decision / outcome / trace は `TaskPaneRefreshOrchestrationService`。execution primitive は `ExcelWindowRecoveryService`。 | display session 内の foreground obligation を terminal outcome にする。 | hidden cleanup、post-close quit、WindowActivate dispatch。 |
| WindowActivate dispatch owner | event capture は `ThisAddIn` / `WorkbookEventCoordinator`。dispatch は `WindowActivatePaneHandlingService`。 | window-safe な TaskPane display / refresh trigger を作る。 | cleanup / recovery / foreground guarantee / CASE display completion の terminal 判定。 |
| trace / outcome owner | 各 protocol owner。display / recovery 系の normalized outcome は `TaskPaneRefreshOrchestrationService`、close / white Excel 系は close lifecycle owner。 | protocol unit ごとに outcome 名と trace owner を分ける。 | lower-level trace を上位 success に丸めること。 |
| user-facing recovery owner | close prompt は `CaseClosePromptService`。folder prompt は `CaseFolderOpenService`。post-close quit failure の UX 正本は未定義。 | 利用者の意思確認と recovery guidance を cleanup protocol から分ける。 | hidden cleanup を user prompt へ委譲すること。不明な UX を実装で補完すること。 |

post-close quit failure や orphaned retained instance detection の user-facing UX は current-state docs では未定義です。target-state では owner を `WindowActivate` や foreground recovery へ移さず、実装前に別安全単位で確認が必要な事項として扱います。

## Owner Boundary Decisions

target-state では、次の境界を固定します。

- `WindowActivate` は cleanup / recovery / foreground guarantee owner ではありません。
- `WindowActivate` は display / refresh trigger に限定します。
- foreground guarantee は hidden Excel cleanup の代替ではありません。
- visibility restore は retained cleanup の代替ではありません。
- `WorkbookClose` は close 後 workbook 再参照を前提にしません。
- retained instance cleanup は user-owned shared app を閉じません。
- isolated app cleanup は isolated owner だけが判断します。
- fail closed を維持します。
- hidden session cleanup failure は、owner 内の catch / poison / release 境界で扱い、別 owner が silent success に丸めません。
- no visible workbook quit は white Excel 防止の close / quit protocol であり、visibility recovery や foreground guarantee ではありません。
- 保存前の workbook window `visible + normal` 正規化は、保存ファイルへ hidden / minimized state を残さないための owner-side cleanup です。shared/current app の display completion ではありません。

## Protocol Target-State

### shared/current app

shared/current app は、Add-in が接続している利用者操作中の Excel `Application` です。

閉じてよい条件:

- `PostCloseFollowUpScheduler` が close 後 follow-up として visible workbook が無いことを確認した場合。
- 対象が current app であり、no visible workbook quit の protocol 範囲内である場合。
- Excel busy retry など既存の close / quit 境界を通り、fail-closed 条件に抵触しない場合。

閉じてはいけない条件:

- visible workbook が 1 つでも残っている場合。
- user-owned shared app か isolated owner か不明な場合。
- `WindowActivate`、foreground guarantee、visibility recovery、hidden cleanup の途中であることだけを理由にする場合。
- workbook / window / context が不明で、推測で補完する必要がある場合。
- close prompt や dirty save 判断が未解決の場合。

### isolated app

isolated app は、処理 owner が生成し、cleanup まで完結させる専用 `Application` です。

閉じてよい条件:

- owner がその isolated app を生成したことが明確である場合。
- owned workbook の work が完了または abort し、owned workbook close / release 境界へ入った場合。
- session owner が shared/current app へ display responsibility を handoff していない isolated 作業 session である場合。
- cleanup owner が `Application.Quit` と COM final release まで実行する場合。

閉じてはいけない条件:

- shared/current app である可能性がある場合。
- retained hidden app-cache の healthy cached app であり、cache owner が keep / idle return と判断する場合。
- owner が生成していない app、または ownership が不明な app の場合。
- foreground guarantee や visibility recovery の失敗を理由にする場合。

### retained hidden app-cache

保持してよい条件:

- `CASEINFO_EXPERIMENT_HIDDEN_APP_CACHE` の例外境界内である場合。
- cached `Application` が healthy で、session close 後に owned workbook が残っていない場合。
- cache owner が idle return 可能と判断し、timeout / poison / shutdown 条件に該当しない場合。

cleanup すべき条件:

- feature flag disabled。
- cached app が poison / unhealthy。
- idle timeout 到達。
- Add-in shutdown。
- session cleanup 失敗により cache owner が retained instance を unsafe と判断した場合。

cleanup してはいけない条件:

- user-owned shared/current app である可能性がある場合。
- cache owner 以外が retained instance を推測で見つけた場合。
- workbook / process の owner tag や cache entry が不明な orphaned `EXCEL.EXE` を、target-state だけで強制終了する場合。

#### retained instance cleanup protocol

retained hidden app-cache は `CaseWorkbookOpenStrategy` が所有する cache entry lifetime です。hidden create session の workbook close と、cached `Application` 自体の cleanup を別 protocol として扱います。

| protocol unit | trigger | owner が出す normalized outcome | runtime mutation | 混ぜてはいけないもの |
| --- | --- | --- | --- | --- |
| keep / return-to-idle | app-cache route の session close で、owned workbook close / release が終わり、cached app が healthy で、timeout / poison / shutdown に該当しない場合。 | `RetainedInstanceReturnedToIdle`。必要に応じて cache 維持を表す `RetainedInstanceKept` と同じ owner vocabulary で扱う。 | cached app を hidden / quiet state に戻し、idle timer を予約する。`Application.Quit` は実行しない。 | isolated app release、retained cleanup completed、foreground guarantee completed。 |
| poison | abort、workbook close failure、return-to-idle failure、health check failure、session cleanup failure で cache owner が reuse unsafe と判断した場合。 | `RetainedInstancePoisoned`。 | cache entry から外し、以後 reuse しない。必要なら slot disposal へ接続する。 | poison を `RetainedInstanceCleanupCompleted` と呼ばない。 |
| timeout cleanup | cached app が idle で、既存 idle timeout に到達した場合。 | `RetainedInstanceCleanupCompleted` / `RetainedInstanceCleanupDegraded` / `RetainedInstanceCleanupFailed`。 | cache-owned app だけ `Application.Quit` / COM release する。 | idle timeout 値変更、in-use app cleanup、shared/current app quit。 |
| feature flag disabled cleanup | cache owner が既存 cache flag disabled を検知した場合。 | `RetainedInstanceCleanupCompleted` / `RetainedInstanceCleanupDegraded` / `RetainedInstancePoisoned`。 | 既存条件内で slot を cache から外す。 | feature flag 意味変更、route 選択条件変更。 |
| shutdown cleanup | Add-in shutdown で `ShutdownHiddenApplicationCache()` が呼ばれた場合。 | `RetainedInstanceCleanupCompleted` / `RetainedInstanceCleanupDegraded` / `RetainedInstanceCleanupSkipped`。 | cached slot と idle timer を cache owner 内で閉じる。 | user-owned shared app cleanup、post-close white Excel quit。 |
| not required / skipped / unknown | cache entry が無い、owner facts が不足、または cache-owned でない slot が見えた場合。 | `RetainedInstanceCleanupNotRequired` / `RetainedInstanceCleanupSkipped` / `RetainedInstanceOwnershipUnknown`。 | cleanup mutation なし、または fail closed。 | process 名だけの broad kill、WindowActivate を根拠にした cleanup。 |

trace policy:

- return-to-idle は session close outcome として記録し、cached app disposal outcome と混同しません。
- timeout / poison / shutdown cleanup は `cleanupReason`、`appQuitAttempted`、`appQuitCompleted` などの raw facts を伴う retained cleanup outcome として記録します。
- hidden cleanup outcome、isolated app outcome、retained instance outcome を同じ success terminal に丸めません。
- retained cleanup は `CaseWorkbookOpenStrategy` cache owner だけが正規化します。`KernelCaseCreationService`、`WindowActivatePaneHandlingService`、`PostCloseFollowUpScheduler` は retained cleanup outcome owner ではありません。
- orphaned `EXCEL.EXE` の監視、通知、強制終了はこの protocol では未定義です。owner facts が揃わない process cleanup は D-3 として扱います。

### hidden Excel visibility restore

visible に戻す条件:

- owner-side save normalization として、owned workbook window を保存前に `visible + normal` へ正規化する場合。
- shared/current app display handoff 後に、presentation / display / recovery owner が対象 workbook window を表示する場合。
- `WorkbookWindowVisibilityService` または `ExcelWindowRecoveryService` が既存条件内で visibility recovery primitive を実行する場合。

戻してはいけない条件:

- retained instance cleanup の代替として行う場合。
- hidden isolated app を利用者表示へ昇格させる場合。
- user-owned hidden window か owner-owned hidden window か不明な場合。
- `WindowActivate` 発火だけを根拠にする場合。

### orphaned instance cleanup

cleanup できる条件:

- cache entry、session facts、owner-owned isolated app である事実が揃っている場合。
- `CaseWorkbookOpenStrategy` の retained cache owner が poison / timeout / shutdown cleanup として判断する場合。
- `KernelUserDataReflectionService` など isolated owner が、自分で生成した app を finally / abort 境界で閉じる場合。

cleanup できない条件:

- PID や process 名だけで user-owned Excel と区別できない場合。
- shared/current app か isolated app か不明な場合。
- `WindowActivate`、foreground recovery、visibility recovery の失敗だけを根拠にする場合。
- target-state docs だけを根拠に broad process kill を導入する場合。

orphaned `EXCEL.EXE` の監視、通知、強制終了の top-level UX / owner は current-state では未定義です。実装前に別安全単位で仕様確認が必要です。

### white Excel recovery

white Excel を recovery 対象とする条件:

- CASE managed close / post-close follow-up 後に visible workbook が無く、Excel だけが残る close / quit protocol 上の状態である場合。
- `PostCloseFollowUpScheduler` が no visible workbook quit の既存条件内で扱う場合。

white Excel を recovery 対象としない条件:

- TaskPane がまだ表示されていないだけの場合。
- foreground に出ていないだけの場合。
- hidden create session / hidden-for-display / retained hidden app-cache の cleanup 問題である場合。
- `WindowActivate` が来ない、または foreground guarantee が degraded / failed であるだけの場合。

### WorkbookClose 後に参照してはいけないもの

`WorkbookClose` 後は、次を再参照してはいけません。

- close 済み workbook COM object。
- close 済み workbook の `Application` / `Windows` / `Sheets` / `CustomDocumentProperties` などの COM member。
- close 済み workbook 由来の active window。
- close 済み workbook に紐づく TaskPane host を、再取得なしに current target とみなすこと。

close 後に使ってよいのは、close 前に採取済みの immutable facts だけです。例として workbook key、path、role、system root、follow-up に必要な owner facts は、close 前に保持した値として扱います。

### reopen 時に再取得すべきもの

reopen 時は、次を再取得します。

- workbook COM object。
- workbook windows / active window。
- workbook role / context / `SYSTEM_ROOT`。
- DocProperty / snapshot / TaskPane context。
- visible pane host 判定。
- foreground / visibility recovery に必要な current window facts。

reopen は、古い closed workbook の延長ではなく、新しい workbook / window facts の取得から始めます。

### foreground guarantee と visibility restore の境界

foreground guarantee:

- display / refresh protocol 内で、対象 Excel / workbook window を前面に戻す obligation を terminal outcome にする unit です。
- decision / outcome / trace owner は `TaskPaneRefreshOrchestrationService` です。
- execution primitive owner は `ExcelWindowRecoveryService` です。

visibility restore:

- owner-owned workbook window を visible に戻す、または保存前に hidden / minimized state を正規化する unit です。
- 保存状態正規化、presentation preparation、lightweight ensure、full recovery primitive に分かれます。

境界:

- visibility restore が成功しても foreground guarantee completed ではありません。
- foreground guarantee が completed / not required でも hidden Excel cleanup completed ではありません。
- `Window.Visible = true`、`workbook.Activate()`、`WindowActivate` 発火はそれぞれ別概念です。

### cleanup と user prompt / UX の境界

- dirty close の user prompt は `CaseClosePromptService` が担当します。
- folder offer は `CaseFolderOpenService` が担当します。
- hidden session cleanup は owner 内で完結させます。利用者 prompt を cleanup completion の条件にしません。
- post-close quit failure や orphaned retained instance detection の user-facing guidance は current-state で未定義です。target-state では実装で補完せず、別安全単位の確認事項として残します。
- cleanup failure を追加 guard、foreground guarantee、visibility restore、`WindowActivate` dispatch で覆いません。

## Normalized Outcome / Trace Policy

この節の outcome 名は target-state の提案です。今回コードには追加しません。

### hidden Excel cleanup

- `HiddenExcelCleanupNotRequired`
- `HiddenExcelCleanupCompleted`
- `HiddenExcelCleanupSkipped`
- `HiddenExcelCleanupDegraded`
- `HiddenExcelCleanupFailed`
- `HiddenExcelCleanupUnknown`

### isolated app lifecycle

- `IsolatedAppReleaseNotRequired`
- `IsolatedAppReleased`
- `IsolatedAppReleaseSkipped`
- `IsolatedAppReleaseDegraded`
- `IsolatedAppReleaseFailed`
- `IsolatedAppOwnershipUnknown`

### retained instance lifecycle

- `RetainedInstanceKept`
- `RetainedInstanceReturnedToIdle`
- `RetainedInstanceCleanupNotRequired`
- `RetainedInstanceCleanupCompleted`
- `RetainedInstanceCleanupSkipped`
- `RetainedInstanceCleanupDegraded`
- `RetainedInstanceCleanupFailed`
- `RetainedInstancePoisoned`
- `RetainedInstanceOwnershipUnknown`

### shared app / white Excel

- `SharedAppQuitNotRequired`
- `SharedAppQuitSkippedVisibleWorkbookExists`
- `SharedAppQuitSkippedOwnershipUnknown`
- `SharedAppQuitCompleted`
- `SharedAppQuitFailed`
- `WhiteExcelRecoveryNotRequired`
- `WhiteExcelRecoveryQueued`
- `WhiteExcelRecoveryCompleted`
- `WhiteExcelRecoverySkipped`
- `WhiteExcelRecoveryDegraded`
- `WhiteExcelRecoveryFailed`

### WorkbookClose / reopen

- `WorkbookCloseNotRequired`
- `WorkbookCloseCompleted`
- `WorkbookCloseSkipped`
- `WorkbookCloseDegraded`
- `WorkbookCloseFailed`
- `WorkbookReopenNotRequired`
- `WorkbookReopenRequired`
- `WorkbookReopenCompleted`
- `WorkbookReopenSkipped`
- `WorkbookReopenFailed`
- `WorkbookReferenceInvalidAfterClose`

### visibility / foreground / WindowActivate

- `VisibilityRestoreNotRequired`
- `VisibilityRestoreCompleted`
- `VisibilityRestoreSkipped`
- `VisibilityRestoreDegraded`
- `VisibilityRestoreFailed`
- `ForegroundGuaranteeNotRequired`
- `ForegroundGuaranteeCompleted`
- `ForegroundGuaranteeDegraded`
- `ForegroundGuaranteeFailed`
- `WindowActivateObserved`
- `WindowActivateDispatchIgnored`
- `WindowActivateDispatchDeferred`
- `WindowActivateDispatchDispatched`
- `WindowActivateDispatchFailed`

trace 方針:

- raw event trace は event capture owner が出します。
- normalized outcome trace は protocol owner が出します。
- cleanup trace と display / foreground trace を同じ success terminal に丸めません。
- `case-display-completed` は display session owner だけが emit します。
- `WindowActivateObserved` は observation であり success terminal ではありません。
- outcome `Unknown` は success completion に使いません。

## Current-State から Target-State への差分

| 現在の混在点 | target-state での owner | 最初に触るべき安全単位 | 後回しにすべき危険単位 | 絶対に触らない条件 |
| --- | --- | --- | --- | --- |
| hidden create session owner が `KernelCaseCreationService` と `CaseWorkbookOpenStrategy` に分かれる。 | session owner は `KernelCaseCreationService`、mechanics / cache owner は `CaseWorkbookOpenStrategy`。 | lifecycle terms / outcome enum docs 固定。 | route 条件や hidden create 条件変更。 | hidden create session の visibility / close 条件。 |
| retained cached app の session close と cached app cleanup が混ざって見える。 | retained instance cleanup owner は `CaseWorkbookOpenStrategy` cache。 | retained instance cleanup protocol の明文化。 | orphaned `EXCEL.EXE` の broad kill。 | user-owned shared app を閉じない。 |
| visibility restore と foreground guarantee が混ざって見える。 | visibility primitive は `WorkbookWindowVisibilityService` / `ExcelWindowRecoveryService`、foreground outcome は `TaskPaneRefreshOrchestrationService`。 | visibility restore と foreground guarantee の境界固定。 | visibility / foreground 条件変更。 | `WindowActivate` を guarantee owner にしない。 |
| white Excel 防止と foreground recovery が混ざって見える。 | white Excel は `PostCloseFollowUpScheduler` の close / quit protocol。 | white Excel recovery の normalized outcome 化。 | post-close quit 条件変更。 | visible workbook がある shared app を quit しない。 |
| WorkbookClose 後の follow-up と reopen が近接する。 | close は lifecycle owner、reopen は presentation / open strategy owner。 | WorkbookClose / reopen protocol の安全化。 | close 後 COM 再参照を伴う修正。 | close 済み workbook を再参照しない。 |
| WindowActivate と activation primitive が混ざって見える。 | WindowActivate dispatch は `WindowActivatePaneHandlingService`、primitive は各 execution owner。 | WindowActivate dispatch outcome の docs / trace 正本化。 | WindowActivate に cleanup / recovery owner を戻すこと。 | WorkbookOpen 直後の window-dependent skip 条件。 |
| cleanup failure UX が未定義。 | user prompt は prompt service、cleanup owner は lifecycle owner。post-close failure UX は未定義。 | user-facing recovery owner の確認事項化。 | UX 名目で cleanup 条件を広げること。 | 不明点を実装で補完しない。 |

最初に着手すべき安全単位は、runtime 条件に触れない `A. lifecycle terms / outcome enum docs 固定` です。次に、既存 trace の意味を変えずに owner ごとの normalized outcome を整理する `B. hidden Excel cleanup owner の trace 正本化` が安全です。

後回しにすべき危険単位は、retained hidden app-cache の実削除、orphaned process cleanup、post-close quit 条件変更、visibility / foreground recovery 条件変更、WorkbookClose / reopen 条件変更です。

絶対に触らない条件:

- `WindowActivate` を cleanup owner に戻すこと。
- DoEvents / sleep / timing hack を提案または追加すること。
- guard 追加で cleanup failure や visibility failure を覆うこと。
- user-owned shared app を quit すること。
- close 済み workbook を再参照すること。
- hidden Excel cleanup 条件、WorkbookClose / reopen 条件、foreground / visibility / rebuild / refresh source 条件を target-state 整理の名目で変更すること。

## 実装安全単位

### A. lifecycle terms / outcome enum docs 固定

目的:

- hidden Excel、isolated app、retained hidden app-cache、white Excel、visibility restore、foreground guarantee、WindowActivate trigger の用語を固定する。
- normalized outcome 名を docs 上で確定し、実装前の合意点にする。

変更してよい範囲:

- docs の用語表、owner 表、outcome 名一覧。
- trace 名の説明。

変更してはいけない範囲:

- コード。
- cleanup 条件。
- visibility / foreground 条件。
- WorkbookClose / reopen 条件。

build / 実機確認ポイント:

- docs-only では build / test / `DeployDebugAddIn` を実行しない。
- 実装に入る場合は `.\build.ps1 -Mode Compile` を標準入口にする。
- runtime `Addins\` 反映や実機確認が必要な場合だけ、別途 `.\build.ps1 -Mode DeployDebugAddIn` を使い、Compile 成功と実機反映成功を分けて扱う。

### B. hidden Excel cleanup owner の trace 正本化

目的:

- hidden session cleanup の owner と trace emit owner を一致させる。
- cleanup success、skipped、degraded、failed、unknown を success に丸めない。

変更してよい範囲:

- hidden cleanup outcome DTO / enum 相当。
- 既存 cleanup owner 内の trace 正規化。
- current condition を変えない範囲の logging。

変更してはいけない範囲:

- hidden workbook open / close 条件。
- isolated app quit 条件。
- shared/current app quit 条件。
- `WindowActivate` dispatch。

build / 実機確認ポイント:

- `.\build.ps1 -Mode Compile`。
- hidden create session の close / abort trace。
- interactive CASE 作成後に hidden session が残らないこと。
- `DeployDebugAddIn` 後の実機確認では runtime assembly trace を確認する。

### C. isolated app lifetime owner の境界固定

目的:

- isolated app を生成した owner だけが `Application.Quit` / COM release を判断する境界を明示する。
- shared/current app fallback と isolated app cleanup を混同しない。

変更してよい範囲:

- `KernelUserDataReflectionService`、`CaseWorkbookOpenStrategy` の owned isolated app outcome / trace。
- ownership facts の受け渡し。

変更してはいけない範囲:

- shared/current app を quit する条件。
- hidden create route の選択条件。
- `AccountingSetKernelSyncService` への isolated fallback 再導入。

build / 実機確認ポイント:

- `.\build.ps1 -Mode Compile`。
- Kernel user data reflection の未 open Base / Accounting 反映。
- interactive CASE 作成 hidden route。
- Excel プロセス残存の観測は owner facts と組み合わせて確認し、process 名だけで判断しない。

### D. retained instance cleanup protocol の明文化

目的:

- retained hidden app-cache の keep / idle return / timeout / poison / shutdown を session close と分ける。
- retained cleanup が user-owned shared app を閉じないことを固定する。

変更してよい範囲:

- cache owner の outcome / trace。
- timeout / poison / shutdown の既存条件の観測強化。

変更してはいけない範囲:

- cache feature flag の意味。
- idle timeout 値。
- cache bypass 条件。
- orphaned `EXCEL.EXE` の broad kill。

build / 実機確認ポイント:

- `.\build.ps1 -Mode Compile`。
- app-cache enabled / disabled の trace。
- cache in-use 時の bypass route。
- Add-in shutdown 時の cache cleanup。

### E. WorkbookClose / reopen protocol の安全化

目的:

- close 後 workbook 再参照を禁止し、reopen は新しい workbook / window facts の再取得として扱う。
- post-close follow-up と immediate reopen の競合を trace できるようにする。

変更してよい範囲:

- close 前 immutable facts の採取。
- close / reopen outcome の trace。
- follow-up queue の diagnostic trace。

変更してはいけない範囲:

- managed close 条件。
- prompt 条件。
- post-close quit 条件。
- reopen 条件。
- close 後 COM member 参照の追加。

build / 実機確認ポイント:

- `.\build.ps1 -Mode Compile`。
- dirty close、clean close、created case folder offer。
- close 後に白 Excel が残らないこと。
- reopen 時に workbook / window / context が再取得されること。

### F. visibility restore と foreground guarantee の境界固定

目的:

- visibility restore、full recovery primitive、foreground guarantee outcome を分ける。
- visibility restore を retained cleanup や foreground success の代替にしない。

変更してよい範囲:

- outcome / trace 名。
- `TaskPaneRefreshOrchestrationService` が受け取る facts の意味付け。
- docs / tests での boundary assertion。

変更してはいけない範囲:

- `WorkbookWindowVisibilityService` / `ExcelWindowRecoveryService` の実行条件。
- foreground guarantee 条件。
- ready-show retry 条件。
- `WindowActivate` 条件。

build / 実機確認ポイント:

- `.\build.ps1 -Mode Compile`。
- CASE 作成直後の ready-show。
- existing CASE reopen。
- foreground guarantee trace と visibility trace が別 outcome として読めること。

### G. white Excel recovery の normalized outcome 化

G-0 の current-state / target boundary は `docs/white-excel-prevention-boundary-current-state.md` で固定します。G-1 で runtime に触る場合も、同文書の安全単位 / 危険単位を前提にします。

目的:

- no visible workbook quit を white Excel prevention protocol として outcome 化する。
- foreground recovery、visibility recovery、WindowActivate と混同しない。

変更してよい範囲:

- `PostCloseFollowUpScheduler` の outcome / trace。
- no visible workbook facts の diagnostic trace。
- quit success / failure / skipped の正規化。

変更してはいけない範囲:

- visible workbook 判定条件。
- `Application.Quit` 実行条件。
- quit 成功後に終了中 app を restore しない既存境界。
- foreground / visibility recovery から post-close quit を呼ぶこと。

build / 実機確認ポイント:

- `.\build.ps1 -Mode Compile`。
- CASE close 後の no visible workbook quit。
- visible workbook が残る場合に shared app を quit しないこと。
- 実機確認では Excel 完全終了、shadow copy、runtime `assemblySha256` を分けて確認する。

## 今回行わないこと

- コード変更なし。
- helper / service 抽出なし。
- Excel visibility 制御変更なし。
- hidden Excel cleanup 条件変更なし。
- WorkbookClose / reopen 条件変更なし。
- foreground / visibility / rebuild / refresh source 条件変更なし。
- `WindowActivate` を cleanup owner に戻す記述なし。
- DoEvents / sleep / timing hack 提案なし。
- guard 追加で覆う方針なし。
- build / test / `DeployDebugAddIn` 実行なし。

## 一言まとめ

target-state では、hidden Excel / isolated app / white Excel lifecycle を「表示が戻ったか」ではなく「どの owner がどの lifecycle を閉じるか」で固定する。shared/current app、isolated app、retained hidden app-cache、WorkbookClose / reopen、foreground guarantee、WindowActivate dispatch はそれぞれ別 unit であり、どれか 1 つの成功を他の cleanup 成功へ読み替えない。
