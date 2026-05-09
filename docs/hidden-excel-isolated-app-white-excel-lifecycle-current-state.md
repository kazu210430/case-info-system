# Hidden Excel / Isolated App / White Excel Lifecycle Current State

## 位置づけ

この文書は、hidden Excel / isolated app / retained hidden app-cache / white Excel lifecycle の current-state 正本です。現行 `main` で確認できる owner、cleanup、visibility、close / reopen 接続点を protocol 単位で固定します。

- 基準コード: `2026-05-08` 作業開始時点で `main` / `origin/main` / `HEAD` が一致した `b9a0f6ad90c607b7fac92b5fbf03f02e90b03390`
- 参照した正本:
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
  - `docs/case-display-recovery-protocol-current-state.md`
  - `docs/case-display-recovery-protocol-target-state.md`
  - `docs/case-workbook-lifecycle-current-state.md`
  - `docs/a2-window-visibility-current-state.md`
  - `docs/workbook-window-activation-notes.md`
  - `docs/taskpane-refresh-policy.md`
- 主な確認対象:
  - `CaseWorkbookOpenStrategy`
  - `KernelCaseCreationService`
  - `KernelCasePresentationService`
  - `KernelUserDataReflectionService`
  - `AccountingSetKernelSyncService`
  - `WorkbookWindowVisibilityService`
  - `ExcelWindowRecoveryService`
  - `TaskPaneRefreshOrchestrationService`
  - `WindowActivatePaneHandlingService`
  - `WorkbookLifecycleCoordinator`
  - `CaseWorkbookLifecycleService`
  - `PostCloseFollowUpScheduler`
  - `ThisAddIn`

この文書は current-state を記録するだけです。コード変更、Excel visibility 制御変更、hidden cleanup 条件変更、WorkbookClose / reopen 条件変更、foreground / visibility / rebuild / refresh source 条件変更は行いません。

## stash 整理結果

作業前に `stash@{0}` を確認しました。

- 対象 stash: `stash@{0}: On codex/windowactivate-owner-boundary: pre-ff-merge debug package artifacts`
- 内容:
  - `dev/Deploy/DebugPackage/CaseInfoSystem.ExcelAddIn/CaseInfoSystem.ExcelAddIn.dll`
  - `dev/Deploy/DebugPackage/CaseInfoSystem.ExcelAddIn/CaseInfoSystem.ExcelAddIn.dll.manifest`
  - `dev/Deploy/DebugPackage/CaseInfoSystem.ExcelAddIn/CaseInfoSystem.ExcelAddIn.vsto`
  - `dev/Deploy/DebugPackage/CaseInfoSystem.WordAddIn/CaseInfoSystem.WordAddIn.dll`
  - `dev/Deploy/DebugPackage/CaseInfoSystem.WordAddIn/CaseInfoSystem.WordAddIn.dll.manifest`
  - `dev/Deploy/DebugPackage/CaseInfoSystem.WordAddIn/CaseInfoSystem.WordAddIn.vsto`
- 判定: `dev/Deploy/DebugPackage` 配下の DLL / manifest / `.vsto` 生成物のみ。
- 実施: stash を drop 済み。drop した stash id は `78747a8e7b9d1856f770a7cfe267307a3dd4c7bb`。

## current-state summary

- 既定は shared/current app です。業務処理側は利用者が操作中の Excel `Application` を原則 quit しません。
- hidden Excel / isolated app は一般的な表示制御手段ではありません。許容されるのは、owner と cleanup が閉じた managed hidden session だけです。
- CASE 新規作成の hidden create session は `KernelCaseCreationService` が session owner、`CaseWorkbookOpenStrategy` が hidden workbook open / close mechanics owner です。
- retained hidden app-cache は `CaseWorkbookOpenStrategy` だけが持つ例外です。workbook close は session close、cached `Application` の return-to-idle / timeout / poison / shutdown cleanup は cache owner が持ちます。
- `KernelUserDataReflectionService` の未 open Base / Accounting 反映は、service-owned isolated app を作り、save 前に owned workbook window visibility を restore し、close / quit / COM final release まで service 内で閉じます。
- `AccountingSetKernelSyncService` は専用 isolated app fallback を持ちません。未 open 会計 workbook も current application で開き、自分で開いた workbook だけ hidden window のまま反映して close します。
- created CASE の interactive 表示は、hidden create session の close 後に shared/current app の hidden-for-display reopen へ移ります。以後の表示、foreground、TaskPane completion は display / refresh protocol 側が扱います。
- white Excel 防止は close / quit 側の設計目標です。現行 owner は `PostCloseFollowUpScheduler` であり、`WindowActivate`、foreground recovery、visibility recovery の代替ではありません。
- `WorkbookActivate` / `WindowActivate` は display / refresh trigger です。hidden session cleanup、retained instance cleanup、white Excel cleanup、foreground guarantee owner ではありません。
- white Excel prevention / recovery の G-0 current-state と target boundary は `docs/white-excel-prevention-boundary-current-state.md` を参照します。

## instance lifecycle 整理

```mermaid
flowchart TD
    Start["Kernel case creation request"] --> Create["KernelCaseCreationService"]
    Create --> HiddenCreate["CaseWorkbookOpenStrategy.OpenHiddenWorkbook"]
    HiddenCreate --> Dedicated["legacy-isolated / experimental-isolated-inner-save / app-cache-bypass-inuse"]
    HiddenCreate --> Cache["app-cache retained hidden application"]
    Dedicated --> DedicatedWork["Initialize -> save -> HiddenCaseWorkbookSession.Close or Abort"]
    DedicatedWork --> DedicatedCleanup["Workbook.Close -> Application.Quit -> COM FinalRelease"]
    Cache --> CacheWork["Initialize -> save -> HiddenCaseWorkbookSession.Close or Abort"]
    CacheWork --> CacheReturn["Workbook.Close -> COM release workbook -> return app to idle"]
    CacheReturn --> CacheCleanup["idle timeout / poison / shutdown -> Application.Quit -> COM FinalRelease"]
    Create --> Batch["CreateCaseBatch"]
    Batch --> BatchEnd["no shared app reopen"]
    Create --> Interactive["NewCaseDefault / CreateCaseSingle"]
    Interactive --> DisplayOpen["OpenHiddenForCaseDisplay on shared/current app"]
    DisplayOpen --> DisplayProtocol["KernelCasePresentationService -> TaskPaneRefreshOrchestrationService"]

    Reflection["KernelUserDataReflectionService unopened Base / Accounting"] --> ReflectionApp["new hidden isolated Application"]
    ReflectionApp --> ReflectionWork["open workbook hidden -> apply -> restore owned window visible -> save"]
    ReflectionWork --> ReflectionCleanup["CloseWorkbookQuietly -> Application.Quit -> COM FinalRelease"]

    Close["CASE WorkbookBeforeClose / managed close"] --> FollowUp["PostCloseFollowUpScheduler"]
    FollowUp --> VisibleCheck["visible workbook check"]
    VisibleCheck --> Quit["no visible workbook -> Application.Quit"]
    VisibleCheck --> Keep["visible workbook exists -> keep Excel"]
```

### shared/current app

shared/current app は、Add-in が接続している利用者操作中の `Application` です。

- created CASE の display reopen は `CaseWorkbookOpenStrategy.OpenHiddenForCaseDisplay(...)` で shared/current app 上に開きます。
- `OpenHiddenForCaseDisplay(...)` は `ScreenUpdating` / `EnableEvents` / `DisplayAlerts` を一時的に false にし、opened workbook window を hidden にして、必要なら previous active window を restore します。
- shared/current app の application state は `RestoreSharedApplicationState(...)` で戻します。
- shared/current app 上で開いた workbook の最終表示責務は `KernelCasePresentationService`、`WorkbookWindowVisibilityService`、`ExcelWindowRecoveryService`、`TaskPaneRefreshOrchestrationService` へ引き継がれます。
- shared/current app を使う経路では、原則として caller-owned `Application` を quit しません。例外は post-close white Excel 防止の no visible workbook quit です。

### isolated app

isolated app は、処理 owner が生成し、cleanup まで完結させる専用 `Application` です。

- `CaseWorkbookOpenStrategy` の dedicated hidden create route:
  - `legacy-isolated`
  - `experimental-isolated-inner-save`
  - `app-cache-bypass-inuse`
- `KernelUserDataReflectionService` の managed hidden reflection session。

isolated app の cleanup 原則:

- owner が `Workbook.Close` を行います。
- owner が `Application.Quit` を行います。
- owner が COM final release を行います。
- close / quit / release を `WindowActivate`、TaskPane refresh、foreground recovery へ委譲しません。

### retained hidden app-cache

retained hidden app-cache は `CaseWorkbookOpenStrategy` の例外境界です。

- 有効化条件は `CASEINFO_EXPERIMENT_HIDDEN_APP_CACHE` です。
- idle timeout は `CASEINFO_EXPERIMENT_HIDDEN_APP_CACHE_IDLE_SECONDS`、未設定時は 15 秒です。
- cache が空なら hidden `Application` を作成して保持します。
- cache が空でなく healthy なら再利用します。
- cache が in-use の場合は `app-cache-bypass-inuse` として dedicated hidden session へ逃がします。
- session close では workbook を close し、workbook COM を release します。
- cached `Application` は healthy なら idle に戻します。
- poison / unhealthy / timeout / feature flag disabled / shutdown では cached `Application` を quit して release します。
- `ThisAddIn_Shutdown` は `CaseWorkbookOpenStrategy.ShutdownHiddenApplicationCache()` を呼びます。

retained hidden app-cache は one-shot isolated lifecycle ではありません。このため orphaned `EXCEL.EXE` の運用監視は残課題ですが、現行 protocol では cache owner 以外が retained instance を破棄しません。

### retained instance cleanup protocol（D-0 current-state）

2026-05-09 の D フェーズ棚卸しでは、開始時 `main` / `origin/main` / `HEAD` が `346b3b33bce887c8de245f42657110483356f7fd` で一致した状態を前提に、`CaseWorkbookOpenStrategy` と `CaseWorkbookOpenStrategyTests` の retained hidden app-cache 周辺だけを確認しました。この節は現行条件の記録であり、cache 有効条件、idle timeout 値、poison 条件、shutdown cleanup 条件、shared/current app quit 条件は変更しません。

| protocol | owner / trigger | current-state action | trace / outcome facts | 境界 |
| --- | --- | --- | --- | --- |
| return-to-idle | `CleanupCachedHiddenSession(...)` が app-cache route の session close を扱い、owned workbook close / release 後に `TryReturnCachedHiddenApplicationToIdle(...)` を呼ぶ。 | cache が有効、slot が同一 `Application`、cache-owned、hidden state reapply と health check が成功した場合だけ `IsInUse=false`、`IdleSinceUtc=DateTime.UtcNow`、idle timer schedule に戻す。`Application.Quit` は行わない。 | `hidden-excel-cleanup-outcome` に `hiddenCleanupOutcome=HiddenExcelCleanupCompleted` と `retainedInstanceOutcome=RetainedInstanceReturnedToIdle`、`cacheReturnedToIdle=True` が出る。 | hidden session cleanup 完了であり、retained cached app cleanup 完了ではない。 |
| poison | abort、workbook close failure、return-to-idle failure、return-to-idle health check failure、open failure cleanup など、cache owner が cached app を再利用不可と判断した場合。 | `MarkCachedHiddenApplicationPoisoned(...)` が slot を cache から外し、idle timer を止め、`DisposeCachedHiddenApplicationSlot(..., "poisoned")` へ渡す。 | `hidden-excel-cleanup-outcome` には `retainedInstanceOutcome=RetainedInstancePoisoned`、`cachePoisoned=True` が出る。slot disposal 側では別に `retained-instance-cleanup-outcome` が出る。 | `RetainedInstancePoisoned` は reuse 禁止の outcome であり、quit / release 完了ではない。 |
| timeout cleanup | idle timer tick または acquire 時の `CleanupExpiredCachedHiddenApplicationUnlocked(...)`。 | cache が null、in-use、または idle timeout 未到達なら disposal しない。poisoned slot または idle timeout 到達 slot だけを cache から外し、timer を止め、dispose へ渡す。 | disposal 側の `retained-instance-cleanup-outcome` に `cleanupReason=idle-timeout` などと `appQuitAttempted` / `appQuitCompleted` が出る。 | timeout 条件は `CASEINFO_EXPERIMENT_HIDDEN_APP_CACHE_IDLE_SECONDS` の既存値に従う。in-use cleanup はしない。 |
| feature-flag-disabled cleanup | idle timer tick で `CASEINFO_EXPERIMENT_HIDDEN_APP_CACHE` が無効化されている場合、または return-to-idle 時に cache disabled が見えた場合。 | timer tick では slot を外して `feature-flag-disabled` reason で dispose する。return-to-idle 中に disabled なら slot を poison 扱いにして return-to-idle しない。 | `retained-instance-cleanup-outcome` の `cleanupReason=feature-flag-disabled`、または poison 経由の outcome で観測する。 | feature flag の意味や route 選択条件は変更しない。 |
| shutdown cleanup | `ThisAddIn_Shutdown` が `CaseWorkbookOpenStrategy.ShutdownHiddenApplicationCache()` を呼ぶ。 | idle timer を dispose し、現在の cached slot だけを cache から外して `shutdown-cleanup` reason で dispose する。 | `retained-instance-cleanup-outcome` に `cleanupReason=shutdown-cleanup` と quit / release の raw facts が出る。 | Add-in shutdown の retained cache cleanup だけを扱い、shared/current app の user-owned workbook / app には広げない。 |
| slot disposal | `DisposeCachedHiddenApplicationSlot(...)`。 | `IsOwnedByCache=false` なら quit せず skip。cache-owned slot だけ `TryQuitApplication(...)` と COM release を行う。 | 現コードで確認できる retained cleanup outcome は `RetainedInstanceCleanupCompleted`、`RetainedInstanceCleanupSkipped`、`RetainedInstanceCleanupDegraded`。`RetainedInstanceCleanupFailed` / `NotRequired` / `OwnershipUnknown` は vocabulary 上の target outcome であり、今回確認範囲では emitted outcome としては未確認。 | process 名や PID だけを根拠にした orphaned `EXCEL.EXE` cleanup は行わない。 |

retained instance cleanup の trace owner は `CaseWorkbookOpenStrategy` です。`hidden-excel-cleanup-outcome` は session close 側の raw facts と retained normalized outcome を併記しますが、cached `Application` を quit / release した事実は `retained-instance-cleanup-outcome` 側で読む必要があります。`WindowActivate`、foreground recovery、visibility restore、white Excel prevention の outcome を retained cleanup success に読み替えません。

## hidden Excel が発生しうる箇所

| 発生箇所 | app 種別 | hidden の意味 | cleanup owner |
| --- | --- | --- | --- |
| CASE 新規作成 hidden create session | isolated / retained | 作成、初期化、保存を画面に出さない作業 session | `KernelCaseCreationService` が session owner、`CaseWorkbookOpenStrategy` が mechanics / cache owner |
| `OpenHiddenForCaseDisplay(...)` | shared/current | shared app 上の CASE reopen を一時的に hidden にして display handoff 前のちらつきと foreground 変化を抑える | shared app state restore は `CaseWorkbookOpenStrategy`。最終表示は presentation / refresh 側 |
| `KernelUserDataReflectionService` 未 open Base / Accounting | isolated | 未 open workbook への反映用 hidden作業 session | `KernelUserDataReflectionService` |
| `AccountingSetKernelSyncService` 未 open accounting workbook | shared/current | current app で自分が開いた workbook window を hidden にして反映する | `AccountingSetKernelSyncService` |
| `MasterWorkbookReadAccessService` / resolver 系の read-only open | shared/current | 読み取り補助のための workbook window hide | 各 read access owner。詳細はこの文書の主対象外 |

## visibility lifecycle 整理

### visibility を変更する主な owner

| owner | current-state の visibility 操作 | protocol 上の意味 |
| --- | --- | --- |
| `CaseWorkbookOpenStrategy.PrepareHiddenApplicationForUse(...)` | hidden app の `Visible=false`、`DisplayAlerts=false`、`ScreenUpdating=false`、`UserControl=false`、`EnableEvents=false` | isolated / retained hidden app の初期状態を固定する。表示回復ではない。 |
| `CaseWorkbookOpenStrategy.HideOpenedWorkbookWindow(...)` | opened workbook の `window.Visible=false` | hidden create / hidden-for-display の作業 window を隠す。 |
| `CaseWorkbookOpenStrategy.RestorePreviousWindow(...)` | previous window の `Visible=true` と `Activate()` | hidden-for-display 中に奪った前景を戻す隣接処理。foreground guarantee ではない。 |
| `KernelCaseCreationService.Normalize*WorkbookWindowStateBeforeSave(...)` | isolated session 内の owned workbook window を `Visible=true`、必要なら `WindowState=xlNormal` | 保存ファイルへ hidden / minimized 状態を残さないための owner-side cleanup。display handoff ではない。 |
| `KernelUserDataReflectionService` | hidden isolated app 作成、target workbook window hide、save 前 window restore | managed hidden reflection session 内の保存状態正規化。shared/current app の表示経路ではない。 |
| `AccountingSetKernelSyncService` | current app quiet scope、owned workbook window hide、owned workbook close | 専用 isolated app fallback ではなく current app 内の owner-owned workbook cleanup。 |
| `WorkbookWindowVisibilityService` | 対象 workbook window を解決し、必要なら `window.Visible=true` | ready-show / presentation 前の lightweight workbook visibility ensure。 |
| `ExcelWindowRecoveryService` | `ScreenUpdating=true`、window visible、window restore、`Application.Visible=true`、`ShowWindow`、`window.Activate()`、foreground promotion | full application / workbook window recovery primitive。 |
| `KernelWorkbookDisplayService` | Kernel HOME 用の Excel / workbook window visibility 制御 | Kernel HOME display / release 境界。CASE display completion owner ではない。 |
| `PostCloseFollowUpScheduler` | visible workbook が無い場合に `DisplayAlerts=false` で `Application.Quit()` | white Excel 防止。visibility recovery ではなく close / quit 側。 |

### foreground / WindowActivate / white Excel recovery との接続

- foreground guarantee:
  - decision / outcome / trace owner は `TaskPaneRefreshOrchestrationService`。
  - execution bridge は `TaskPaneRefreshCoordinator`。
  - execution primitive は `ExcelWindowRecoveryService`。
- `WindowActivate`:
  - event capture は `ThisAddIn`。
  - request 化と dispatch は `WindowActivatePaneHandlingService`。
  - refresh protocol へ到達した後の outcome 正規化は `TaskPaneRefreshOrchestrationService`。
  - `WindowActivateDispatchOutcome` は display completion、recovery owner、foreground guarantee owner、hidden Excel owner のいずれでもありません。
- white Excel:
  - CASE close 後に visible workbook が無い場合の quit 判定は `PostCloseFollowUpScheduler`。
  - `WindowActivate`、foreground guarantee、visibility recovery を post-close quit の代替 owner としません。
  - `targetWorkbookStillOpen` と `visibleWorkbookExists` の current meaning は `docs/white-excel-prevention-boundary-current-state.md` で固定します。

## WorkbookClose / reopen との接続点

WorkbookClose / reopen protocol の E-0 詳細 current mapping は `docs/workbook-close-reopen-protocol-current-mapping.md` を参照します。

### WorkbookClose

- Excel event は `ThisAddIn.Application_WorkbookBeforeClose(...)` で受けます。
- `WorkbookLifecycleCoordinator.OnWorkbookBeforeClose(...)` が CASE / Kernel / Accounting lifecycle service へ順に渡します。
- CASE dirty path では `CaseWorkbookLifecycleService` が prompt、folder offer、managed close、post-close follow-up を調停します。
- managed close では `ManagedCloseState` scope 内で save 有無を扱い、対象 workbook close 後に post-close follow-up を予約します。
- white Excel 防止は post-close follow-up 側で visible workbook を確認し、無ければ quit します。

### reopen

- interactive created CASE は hidden create session close 後に `KernelCasePresentationService.OpenCreatedCase(...)` から shared/current app へ reopen します。
- `NewCaseDefault` / `CreateCaseSingle` では `OpenHiddenForCaseDisplay(...)` が選ばれます。
- reopen 直後は workbook window を一時 hidden にして previous window を戻し、その後 `KernelCasePresentationService.ShowCreatedCase(...)` が visibility ensure、without-showing recovery、ready-show request へ進めます。
- `WorkbookOpen` は window 安定境界ではありません。window-dependent refresh は `WorkbookActivate` / `WindowActivate` 以降、または ready-show / retry 側で扱います。

## owner 分裂 / 混在ポイント

- CASE hidden create session の owner が分かれています。
  - session owner は `KernelCaseCreationService`。
  - hidden workbook open / close mechanics は `CaseWorkbookOpenStrategy`。
  - retained cached `Application` owner は `CaseWorkbookOpenStrategy`。
- visibility recovery が分かれています。
  - lightweight workbook window visible ensure は `WorkbookWindowVisibilityService`。
  - full app/window/foreground recovery primitive は `ExcelWindowRecoveryService`。
  - foreground guarantee decision / outcome / trace は `TaskPaneRefreshOrchestrationService`。
  - Kernel HOME visibility は `KernelWorkbookDisplayService`。
- hidden-for-display と display completion が分かれています。
  - `OpenHiddenForCaseDisplay(...)` は shared app reopen と一時 hidden / previous window restore まで。
  - `case-display-completed` は `TaskPaneRefreshOrchestrationService` の created-case display session terminal。
- `WorkbookActivate` / `WindowActivate` と activation primitive が混在して見えます。
  - event trigger は lifecycle / dispatch owner。
  - `workbook.Activate()`、`window.Activate()`、`ShowWindow`、`SetForegroundWindow` は各 primitive owner。
- white Excel 防止と foreground recovery が混在して見えます。
  - white Excel 防止は close / quit 側の `PostCloseFollowUpScheduler`。
  - foreground recovery は display / refresh 側の `TaskPaneRefreshOrchestrationService` と `ExcelWindowRecoveryService`。
- `Application` state restore の範囲が複数あります。
  - hidden-for-display shared app state restore は `CaseWorkbookOpenStrategy`。
  - reflection quiet mode restore は `KernelUserDataReflectionService`。
  - accounting sync quiet scope restore は `ExcelApplicationStateScope`。
  - post-close quit 成功後は終了中 application を restore しません。

## protocol 上の未定義ポイント

current-state では次を未定義または暗黙の protocol として扱います。

- retained hidden app-cache の運用上の必要性。
- retained hidden app-cache に起因する orphaned `EXCEL.EXE` を検出、通知、強制終了する top-level owner。
- cached `Application` が idle timeout 前に外部要因で不健康になった場合の利用者向け recovery 表示。
- `PostCloseFollowUpScheduler` の `Application.Quit()` 失敗後に、どの UX / retry / manual guidance を正本とするか。
- user が WorkbookClose 直後に同じ CASE を即 reopen した場合の post-close follow-up queue との衝突扱い。
- Kernel HOME visibility owner と CASE foreground guarantee owner が同時期に走る場合の protocol 名。
- `WorkbookActivate` と `WindowActivate` のどちらをすべての環境で最終安全境界にするべきか。
- read-only / temporary workbook close 全般を hidden lifecycle 正本へ含めるかどうか。
- hidden window state を保存ファイルへ残さないための normalization を、CASE create 以外の workbook 種別へ一般化するかどうか。
- white Excel という運用呼称と、`no visible workbook -> quit` protocol の正式名称。

不明な事項は current-state では補完しません。target-state 化まで条件変更や guard 追加で覆わない前提です。

## fail closed 境界

- workbook / window / context が不明な場合、`WindowActivate` や foreground recovery を根拠に推測で補完しません。
- `WorkbookOpen` 直後に window が未確定な refresh は shared policy で skip し、後続の `WorkbookActivate` / `WindowActivate` / ready-show / retry 側へ委ねます。
- hidden session cleanup が失敗した場合、owner 内の catch / poison / release 境界で扱います。別 owner が silent success に丸める protocol ではありません。
- `PostCloseFollowUpScheduler` は visible workbook がある場合 quit しません。visible 判定を bypass して white Excel 対策を広げる定義はありません。
- `Application.DoEvents()`、sleep、timing hack、追加 foreground guard で不明点を隠す方針は採りません。

## 守るべき既存制約

- 白Excel対策を壊さない。
- TaskPane 不表示 regression を防ぐ。
- hidden create session、hidden-for-display、managed hidden reflection session、retained hidden app-cache の owner / cleanup 境界を広げない。
- `ScreenUpdating` / `DisplayAlerts` / `EnableEvents` を変更した場合は既存 scope で復元する。ただし `Application.Quit()` 成功後の終了中 app は restore しない。
- COM release を落とさない。
- `WorkbookOpen` を window 安定境界として扱わない。
- `WindowActivate` を recovery / guarantee / hidden cleanup owner にしない。
- `WorkbookClose` / reopen 条件を変えない。
- foreground / visibility / rebuild fallback / refresh source 条件を変えない。
- service 分割、helper 切り出し、context-less workbook 推測、暗黙 open を追加しない。

## 次に target-state 化すべき論点

1. retained hidden app-cache を今後も維持するか、実運用上の必要性と orphaned `EXCEL.EXE` 監視方針を決める。
2. hidden create session owner を `KernelCaseCreationService` と `CaseWorkbookOpenStrategy` に分けたまま、protocol 名だけをどう固定するか。
3. `OpenHiddenForCaseDisplay(...)` を shared/current app の display handoff 前処理としてどこまで target-state に含めるか。
4. white Excel prevention を `PostCloseFollowUpScheduler` の close / quit protocol として独立 target-state 化するか。
5. `WorkbookClose -> post-close follow-up -> no visible workbook quit` と `reopen` の競合条件をどう観測、記録、fail closed 化するか。
6. visibility lifecycle を `workbook window visible ensure`、`application/window recovery`、`foreground guarantee`、`post-close quit` の 4 層で命名固定するか。
7. Kernel HOME visibility と CASE display visibility の接続点を同じ target-state に含めるか、別文書で扱うか。
8. `WindowActivate` target-state と hidden / white Excel lifecycle target-state の参照関係を、相互除外として固定するか。
9. read-only / temporary workbook close の helper 非経由箇所を hidden lifecycle の対象に含めるか。
10. target-state 化後も docs-only の検証観点として、build 成功と runtime `Addins\` 反映成功を混同しない運用を維持する。

## 今回行わないこと

- コード変更なし。
- Excel visibility 制御変更なし。
- hidden Excel cleanup 条件変更なし。
- WorkbookClose / reopen 条件変更なし。
- foreground / visibility / rebuild / refresh source 条件変更なし。
- service 分割なし。
- helper 切り出しなし。
- build / test / `DeployDebugAddIn` 実行なし。docs-only 指示のため実行しない。
