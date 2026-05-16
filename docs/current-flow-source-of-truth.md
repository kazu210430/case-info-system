# Current Flow Source of Truth

## 位置づけ

この文書は、案件情報System の現行挙動を `main` 基準で固定するための横断正本です。目的は「理想構造を書くこと」ではなく、実際に今どう動いているか、どこで fail-closed しているか、どこに ownership 混在が残っているかを将来の安全な最小切り出しの前提として保存することです。

- 参照前提:
  - `AGENTS.md`
  - `docs/architecture.md`
  - `docs/flows.md`
  - `docs/ui-policy.md`
  - `docs/taskpane-architecture.md`
  - `docs/taskpane-refresh-policy.md`
  - `docs/workbook-window-activation-notes.md`
  - `docs/thisaddin-boundary-inventory.md`
  - `docs/*current-state*.md`
  - `docs/*responsibility*.md`

## 読み方

- source-of-truth は「現行 `main` の実コード順序と、それを固定済みの current-state / responsibility docs」です。
- `WorkbookOpen` と window-safe 境界を混同しません。
- `WorkbookContext` と `SYSTEM_ROOT` が明示されている経路では、その文脈を唯一の入口として扱います。
- hidden session は一般解ではなく、owner と cleanup が閉じた例外だけを認めます。
- docs に根拠がない設計意図は断定せず、「不明」として残します。

## 1. startup / shutdown

### フロー概要

- 入口:
  - `ThisAddIn_Startup(...)`
  - `ThisAddIn_Shutdown(...)`
- source-of-truth:
  - `dev/CaseInfoSystem.ExcelAddIn/ThisAddIn.cs`
  - `docs/thisaddin-boundary-inventory.md`
- context:
  - `AddInCompositionRoot`
  - `ApplicationEventSubscriptionService`
  - startup 時の `KernelStartupContext`
- fail-closed 条件:
  - startup context が Kernel HOME 自動表示条件を満たさない場合は HOME を出さない
  - shutdown は cleanup を優先し、表示回復や再入を試みない
- 不明:
  - startup 文脈で `WorkbookActivate` と `WindowActivate` のどちらを最終安全境界とみなすべきかは、コードだけでは確定しない

### 時系列フロー

1. `InitializeStartupDiagnostics()` が logger と起動 trace を初期化する。
2. `CreateStartupCompositionRoot() -> Compose() -> ApplyCompositionRoot()` で service 群を組み立てる。
3. `InitializeApplicationEventSubscriptionService()` が Excel event 購読対象を構成する。
4. `HookApplicationEvents()` が `WorkbookOpen`、`WorkbookActivate`、`WorkbookBeforeSave`、`WorkbookBeforeClose`、`WindowActivate`、`SheetActivate`、`SheetSelectionChange`、`SheetChange`、`AfterCalculate` を購読する。
5. `TryShowKernelHomeFormOnStartup()` が `KernelWorkbookService.ShouldShowHomeOnStartup()` を通して startup context を判定する。
6. `shouldShow == true` のときだけ HOME binding をクリアし、`ShowKernelHomePlaceholder()` へ進む。
7. startup の最後に `RefreshTaskPane("Startup", null, null)` を呼ぶ。
8. shutdown では `UnhookApplicationEvents()`、pending pane refresh timer 停止、HOME form close、`TaskPaneManager.DisposeAll()`、Word warm-up timer 停止、`CaseWorkbookOpenStrategy.ShutdownHiddenApplicationCache()` の順で cleanup する。

### 責務境界

- VSTO lifecycle:
  - `ThisAddIn`
- service composition:
  - `AddInCompositionRoot`
- startup 事実収集:
  - `KernelStartupContextInspector`
- startup 表示判定:
  - `KernelWorkbookStartupDisplayPolicy`
- 初回 HOME 表示:
  - `ThisAddIn` + `KernelWorkbookService`
- 初回 TaskPane refresh:
  - `TaskPaneRefreshOrchestrationService`
- shutdown cleanup:
  - `ThisAddIn`、`TaskPaneManager`、`CaseWorkbookOpenStrategy`

### 現在の ownership 混在

- `ThisAddIn` が VSTO lifecycle、event handler、HOME form instance、TaskPane display entry、`CustomTaskPane` create/remove adapter、automation public surface を同時保持している。
- startup/shutdown 順序、event wiring、TaskPane display adapter が add-in 境界に近接しており、読みやすさ改善と ownership 分離がまだ一致していない。
- `ThisAddIn` から各 coordinator へ委譲していても、最終的な VSTO 境界と UI 境界は `ThisAddIn` に残る。

### 危険領域

- ordering-sensitive
- WorkbookOpen-sensitive
- runtime-sensitive
- window-sensitive
- hidden-session-sensitive

### 将来の切り出し候補

- startup pipeline service
- shutdown pipeline service
- application event wiring adapter
- HOME form host / lifecycle adapter

## 2. Kernel HOME

### フロー概要

- 入口:
  - `TryShowKernelHomeFormOnStartup()`
  - `ShowKernelHomeFromAutomation()`
  - `ShowKernelHomeFromKernelCommand()`
  - `ShowKernelSheetAndRefreshPaneFromHome(...)`
- source-of-truth:
  - `ThisAddIn.ShowKernelHomePlaceholder(...)`
  - `KernelWorkbookService`
  - `docs/ui-policy.md`
  - `docs/flows.md`
- context:
  - valid binding がある bound HOME
  - binding を持たない `unbound` HOME
- fail-closed 条件:
  - `unbound` HOME は placeholder-only とし、Kernel workbook の探索・open・自動 bind をしない
  - bound 文脈が解決できない場合、sheet 遷移や refresh は実行しない
- 不明:
  - HOME 表示不整合時の期待 UX の全件はコードだけでは確定しない

### 時系列フロー

1. startup または明示操作が `ShowKernelHomePlaceholderWithExternalWorkbookSuppressionCore(...)` もしくは `TryShowKernelHomeFormOnStartup()` に入る。
2. 明示表示経路では `SuppressUpcomingKernelHomeDisplay(..., suppressOnOpen: false, suppressOnActivate: true)` を発行する。
3. 必要なら既存の非表示 `KernelHomeForm` を silent dispose し、新セッション時は HOME binding をクリアしてから form を再生成する。
4. `TaskPaneManager.HideKernelPanes()` で Kernel pane を隠し、`KernelHomeForm.ReloadSettings()` の後に `PrepareForHomeDisplayFromSheet()` と `EnsureHomeDisplayHidden(...)` を呼ぶ。
5. `KernelHomeForm.Show() -> Activate() -> BringToFront()` で独立フォームとして表示する。
6. HOME からの sheet 遷移では `ShowKernelSheetAndRefreshPaneFromHome(...)` が `WorkbookContext` を必須とし、`ResolveKernelWorkbook(context)` に失敗したら abort する。
7. 遷移前に `SuppressUpcomingKernelHomeDisplay(...)` を入れ、必要時だけ `RunWithScreenUpdatingSuspended(...)` を使い、`HideKernelHomePlaceholder() -> TryShowSheetByCodeName(...) -> RefreshTaskPane(...)` の順で遷移する。

### 責務境界

- HOME form UI:
  - `KernelHomeForm`
- HOME binding / root 整合:
  - `KernelWorkbookBindingService`
- HOME 表示準備 / visibility:
  - `KernelWorkbookDisplayService`
- HOME close backend:
  - `KernelWorkbookCloseService`
- HOME suppression:
  - `KernelHomeCasePaneSuppressionCoordinator`
- sheet 遷移:
  - `ThisAddIn` + `KernelWorkbookService`

### 現在の ownership 混在

- `ThisAddIn` が HOME form instance の生成・破棄・表示と suppression 発行を持ち、`KernelWorkbookService` が binding/display/close facade を持つため、HOME は UI owner と backend owner が分かれている。
- `ShowKernelHomePlaceholder(...)` の中に form lifecycle、TaskPane hide、display preparation、WinForms 表示が同居している。
- HOME suppression と CASE pane suppression が同じ coordinator に同居している。

### 危険領域

- fail-closed-sensitive
- ordering-sensitive
- ScreenUpdating-sensitive
- window-sensitive
- external-workbook-sensitive

### 将来の切り出し候補

- HOME form host service
- HOME navigation coordinator
- HOME suppression state holder
- HOME display-entry adapter

## 3. CASE 作成

### フロー概要

- 入口:
  - `KernelCaseCreationService`
  - `KernelCaseCreationCommandService`
- source-of-truth:
  - `docs/flows.md`
  - `docs/architecture.md`
- context:
  - `SYSTEM_ROOT`
  - `NAME_RULE_A`
  - `NAME_RULE_B`
  - Base workbook path
  - creation mode (`NewCaseDefault` / `CreateCaseSingle` / `CreateCaseBatch`)
- fail-closed 条件:
  - root、Base、出力先、初期化対象が揃わなければ作成を進めない
  - interactive handoff 前に hidden create session を閉じ切れない場合、表示経路へ昇格しない
- 不明:
  - `CaseWorkbookInitializer` が初期化時に書き込む全項目一覧は、この文書だけでは確定しない

### 時系列フロー

1. `KernelCaseCreationService` が `SYSTEM_ROOT`、name rule、Base path、出力先フォルダ、CASE 名を解決する。
2. Base を物理コピーして CASE workbook を作成する。
3. 現行 `main` では `ShouldUseHiddenCreateSession() == true` のため、全モードが hidden create route を通る。
4. `CaseWorkbookOpenStrategy.OpenHiddenWorkbook(...)` が `app-cache`、`legacy-isolated`、`experimental-isolated-inner-save` のいずれかを選ぶ。
5. interactive route (`NewCaseDefault` / `CreateCaseSingle`) は hidden create session 内で visible create 初期化を行うが、save 前に owned workbook window を `Visible=true` へ戻さない。`NormalizeInteractiveWorkbookWindowStateBeforeSave(...)` は `save-window-visible-deferred` として表示責務を `KernelCasePresentationService` へ遅延し、必要なら `WindowState=xlNormal` だけを整えて save / hidden session close を完了する。
6. batch route (`CreateCaseBatch`) は save 前に workbook window を `visible + normal` へ正規化するが、表示経路へ昇格させず reopen もしない。
7. interactive route では hidden session close 後に `KernelCasePresentationService` が shared/current app 側の reopen と表示責務を引き継ぐ。
8. CASE 作成中の Kernel HOME close は `KernelWorkbookCloseService` / `KernelHomeSessionDisplayPolicy` が display restore を skip し、CASE より前に Kernel を戻さない。

### 責務境界

- creation plan / path resolve:
  - `KernelCaseCreationService`
- hidden workbook open / retained app-cache:
  - `CaseWorkbookOpenStrategy`
- workbook 初期化:
  - `CaseWorkbookInitializer`
- interactive 表示 handoff:
  - `KernelCasePresentationService`
- HOME close side effect:
  - `KernelWorkbookCloseService`

### 現在の ownership 混在

- CASE 作成 1 本の業務フローが、creation owner、hidden open owner、retained cache owner、interactive 表示 owner、HOME close owner にまたがる。
- hidden create session の owner は `KernelCaseCreationService` だが、open/close mechanics は `CaseWorkbookOpenStrategy` にあり、interactive handoff は `KernelCasePresentationService` に移る。
- creation 自体は非表示作業であり、保存状態正規化の扱いは interactive / batch で分かれる。interactive は `Visible=true` を戻さず `WindowState=xlNormal` だけを許容し、batch は save 前に `visible + normal` へ正規化する。

### 危険領域

- hidden-session-sensitive
- ordering-sensitive
- window-sensitive
- COM-sensitive
- fail-closed-sensitive

### 将来の切り出し候補

- create plan builder
- hidden create session owner abstraction
- interactive handoff coordinator
- retained hidden app-cache inventory boundary

## 4. CASE 表示

### フロー概要

- 入口:
  - `KernelCasePresentationService.OpenCreatedCase(...)`
- source-of-truth:
  - `KernelCasePresentationService`
  - `docs/flows.md`
  - `docs/taskpane-refresh-policy.md`
- context:
  - `KernelCaseCreationResult`
  - CASE workbook path
  - wait UI session
  - CASE creation mode
- fail-closed 条件:
  - `result.Success == false` なら表示へ進まない
  - path や workbook が解決できない場合は例外で止め、wait UI close と suppression release を優先する
- 不明:
  - final foreground 安定化の UX 完了条件は実機観測なしでは断定しない

### 時系列フロー

1. `OpenCreatedCase(...)` が result と path を検証し、必要なら待機 UI を開く。
2. `RegisterKnownCasePath(...)` と `TransientPaneSuppressionService.SuppressPath(...)` を入れる。
3. `OpenCreatedCaseWorkbook(...)` が interactive mode では `OpenHiddenForCaseDisplay(...)`、それ以外では visible open を選ぶ。
4. `ShowCreatedCase(...)` が `WorkbookWindowVisibilityService.EnsureVisible(...)` を呼び、shared app visibility recovery 前の window 可視化補助を行う。
5. `ExcelWindowRecoveryService.TryRecoverWorkbookWindowWithoutShowing(..., bringToFront: false)` で初期 recovery を行う。
6. deferred presentation で `TransientPaneSuppressionService.ReleaseWorkbook(...)` を実行する。
7. その直後に再度 `EnsureVisible(...)` を行い、`SuppressUpcomingCasePaneActivationRefresh(...)` を設定する。
8. `ShowWorkbookTaskPaneWhenReady(...)` で ready-show を要求する。
9. `MoveInitialCursorToHomeCell(...)` が HOME シートを activate し、必要なら一時 read-only open した Kernel から初期カーソル位置を解いて `Range.Select()` 後に `FinalRelease(range)` する。
10. wait UI を close し、`NewCaseDefault` 以外では `PromoteWorkbookWindowOnce(...)` が Excel hwnd と workbook hwnd を一度だけ前面化する。

### 責務境界

- reopen / open strategy:
  - `CaseWorkbookOpenStrategy`
- wait UI:
  - `CreatedCasePresentationWaitService`
- transient suppression:
  - `TransientPaneSuppressionService`
- workbook visibility ensure:
  - `WorkbookWindowVisibilityService`
- Excel window recovery:
  - `ExcelWindowRecoveryService`
- ready-show handoff:
  - `TaskPaneRefreshOrchestrationService`
- initial cursor resolve:
  - `KernelCasePresentationService` + `KernelWorkbookResolverService`

### 現在の ownership 混在

- `KernelCasePresentationService` が wait UI、suppression、visibility recovery、ready-show handoff、cursor positioning、最終 foreground promotion をまとめて持つ。
- interactive 表示本線が、open strategy、window recovery、TaskPane ready-show、cursor positioning にまたがる orchestrator になっている。
- CASE 表示完了の判定が workbook visible、pane ready-show、cursor 移動、wait UI close にまたがり、1 つの単純な completion state に閉じていない。

### 危険領域

- runtime-sensitive
- window-sensitive
- foreground-sensitive
- hidden-session-sensitive
- ordering-sensitive

### 将来の切り出し候補

- created-case wait UI coordinator
- visibility recovery handoff service
- initial cursor resolver
- final foreground promotion helper

## 5. TaskPane lifecycle

### フロー概要

- 入口:
  - `WorkbookLifecycleCoordinator` (`WorkbookOpen` / `WorkbookActivate`)
  - `WindowActivatePaneHandlingService` (`WindowActivate`)
  - `TaskPaneRefreshOrchestrationService` (explicit refresh / ready-show)
- source-of-truth:
  - `TaskPaneRefreshPreconditionPolicy`
  - `TaskPaneRefreshOrchestrationService`
  - `TaskPaneManager`
  - `docs/taskpane-architecture.md`
  - `docs/taskpane-refresh-policy.md`
- context:
  - `WorkbookContext`
  - workbook / window
  - window key
  - CASE cache / Base cache / Master rebuild
- fail-closed 条件:
  - `WorkbookOpen` 直後に window が無ければ refresh 完了にしない
  - precondition が false なら refresh せず skip / defer に回す
- 不明:
  - retry 秒数や attempts の正式な業務仕様根拠は docs だけでは確定しない

### 時系列フロー

1. event または明示要求が `ThisAddIn.RefreshTaskPane(...)` か `ShowWorkbookTaskPaneWhenReady(...)` に入る。
2. 通常 refresh では `TaskPaneRefreshOrchestrationService.TryRefreshTaskPane(...)` が `RefreshPreconditionEvaluator` を通し、protection・suppression・`WorkbookOpen` skip を先に判定する。
3. `WorkbookOpen` 直後かつ `workbook != null && window == null` なら `TaskPaneRefreshPreconditionPolicy.ShouldSkipWorkbookOpenWindowDependentRefresh(...)` により skip する。
4. refresh が許可された場合、`TaskPaneRefreshCoordinator` が `WorkbookContext` を解決して `TaskPaneManager.RefreshPane(...)` に渡す。
5. `TaskPaneManager` / `TaskPaneHostFlowService` が hide / skip / stale cleanup / host reuse / render / show を調停する。host-flow entry の `hide-all / skip` reason 判定自体は `TaskPaneRefreshPreconditionPolicy` を正本として消費する。
6. CASE 表示直後の ready-show は `WorkbookTaskPaneReadyShowAttemptWorker.ShowWhenReady(...)` が担当し、attempt 1 でだけ `WorkbookWindowVisibilityService.EnsureVisible(...)` を実行する。
7. worker は `ResolveWorkbookPaneWindow(...)` で window を解決し、visible CASE pane が既にあれば early-complete で成功相当終了する。
8. early-complete しない場合だけ `TryRefreshTaskPane(...)` へ handoff する。
9. ready-show は `80ms` の retry を最大 2 attempt 行い、尽きた場合だけ `ScheduleWorkbookTaskPaneRefresh(...)` から `400ms` pending retry 最大 3 attempt に handoff する。
10. pending retry は対象 workbook を追い、見失っても active CASE context が残る場合は active refresh fallback を継続する。
11. 既存 CASE の Pane / host / control は close まで維持し、`WorkbookActivate` / `WindowActivate` のたびに version 比較で再生成しない。

### 責務境界

- event-side refresh orchestration:
  - `TaskPaneRefreshOrchestrationService`
- window resolve:
  - `WorkbookPaneWindowResolver`
- ready-show attempt:
  - `WorkbookTaskPaneReadyShowAttemptWorker`
- workbook visible ensure:
  - `WorkbookWindowVisibilityService`
- pending retry:
  - `PendingPaneRefreshRetryService`
- host flow / render / show:
  - `TaskPaneManager` / `TaskPaneHostFlowService` / `TaskPaneDisplayCoordinator`
- snapshot source choose:
  - `TaskPaneSnapshotBuilderService`

### 現在の ownership 混在

- `TaskPaneManager` は facade でありながら host map owner と facade entry surface を兼ねる。
- `TaskPaneHostFactory` が CASE / Kernel / Accounting control の event 配線をまとめて持ち、handler 分離と wiring owner 分離が一致していない。
- `TaskPaneActionDispatcher` の post-action refresh が `ThisAddIn.RequestTaskPaneDisplayForTargetWindow(...)` に再入し、CASE action dispatch と display entry が完全には分離していない。
- `CasePaneCacheRefreshNotificationService` が notification だけでなく `WorkbookOpen` / `WorkbookActivate` timing と workbook `Saved` restore を持つ。

### 危険領域

- WorkbookOpen-sensitive
- visibility/foreground-sensitive
- runtime-sensitive
- UX-sensitive
- host-metadata-sensitive

### 将来の切り出し候補

- shared host state holder の明文化
- dispatcher subtree の composition owner 整理
- `TaskPaneHostRegistry` / `TaskPaneHostFactory` の VSTO boundary inventory 分離
- notification service の lifecycle adjacency 切り出し

## 6. window activation / foreground stabilization

### フロー概要

- 入口:
  - `Application_WindowActivate(...)`
  - CASE ready-show 後の foreground guarantee
  - CASE 表示後の `PromoteWorkbookWindowOnce(...)`
- source-of-truth:
  - `WindowActivatePaneHandlingService`
  - `KernelHomeCasePaneSuppressionCoordinator`
  - `ExcelWindowRecoveryService`
  - `docs/workbook-window-activation-notes.md`
  - `docs/a2-window-visibility-current-state.md`
- context:
  - workbook / window
  - active workbook / active window
  - suppression / protection state
- fail-closed 条件:
  - protection 中は `WorkbookActivate` / `WindowActivate` / `TaskPaneRefresh` を無理に通さず return する
  - suppression 中は CASE pane refresh を進めない
- 不明:
  - protection 秒数 `5秒` の正式な仕様根拠は docs だけでは確定しない

### 時系列フロー

1. `Application_WindowActivate(...)` が trace と active state logging を行い、`WorkbookEventCoordinator.OnWindowActivate(...)` へ渡す。
2. `HandleWindowActivateEvent(...)` が `WindowActivatePaneHandlingService.Handle(...)` を呼ぶ。
3. service は `TaskPaneDisplayRequest.ForWindowActivate()` を作り、最初に `ShouldIgnoreDuringCaseProtection(...)` を判定する。
4. protection で止まらなければ `HandleExternalWorkbookDetected(...)` を実行する。
5. その後 `ShouldSuppressCasePaneRefresh(...)` が真なら suppression return する。
6. suppression にも当たらなければ `RequestTaskPaneDisplayForTargetWindow(...)` を通して refresh へ進む。
7. CASE 表示の deferred presentation では、`ReleaseWorkbook(...) -> EnsureVisible(...) -> SuppressUpcomingCasePaneActivationRefresh(...) -> ShowWorkbookTaskPaneWhenReady(...)` の順で activation refresh suppression を設定する。
8. CASE refresh 成功後は `BeginCaseWorkbookActivateProtection(...)` が入り、`WorkbookActivate` / `WindowActivate` / `TaskPaneRefresh` の 3 入口で active window 基準の protection が効く。
9. window placement current-state では、snapped CASE window に対する `ShowWindow(SW_RESTORE)` の normalize side effect が危険点として観測されており、visible + `SW_SHOWNORMAL` + not minimized/maximized では restore を skip する current-state を維持する。

### 責務境界

- WindowActivate entry:
  - `ThisAddIn` / `WindowActivatePaneHandlingService`
- suppression / protection state:
  - `KernelHomeCasePaneSuppressionCoordinator`
- workbook visible ensure:
  - `WorkbookWindowVisibilityService`
- foreground recovery primitive:
  - `ExcelWindowRecoveryService`
- CASE ready-show downstream:
  - `TaskPaneRefreshCoordinator`
- one-shot promotion:
  - `KernelCasePresentationService`

### 現在の ownership 混在

- 同一 coordinator が Kernel HOME suppression、CASE pane suppression、CASE foreground protection を同時保持している。
- `WindowActivatePaneHandlingService` は protection predicate、external workbook 検知、suppression 判定、refresh 入口を直列で持つ。
- foreground stabilization が ready-show、window recovery、TaskPane refresh 成功、WindowActivate 側 suppression と連動している。

### 危険領域

- window-sensitive
- foreground-sensitive
- ordering-sensitive
- runtime-sensitive
- placement-sensitive

### 将来の切り出し候補

- visible window resolve ownership の単独棚卸し
- foreground retry semantics inventory
- protection state owner の独立化
- placement restore decision inventory

## 7. Workbook close / COM release timing

### フロー概要

- 入口:
  - `KernelHomeForm.FormClosing`
  - `WorkbookBeforeClose`
  - hidden session cleanup
  - temporary COM object cleanup
- source-of-truth:
  - `CaseWorkbookLifecycleService`
  - `KernelWorkbookCloseService`
  - `KernelUserDataReflectionService`
  - `docs/case-workbook-lifecycle-current-state.md`
  - `docs/accounting-close-lifecycle-current-state.md`
- context:
  - workbook role
  - dirty state
  - managed close state
  - owned workbook / owned application か shared か
- fail-closed 条件:
  - HOME close は backend close 成功前に Form を閉じない
  - close 後の対象 workbook 再参照をしない
  - cleanup failure が元例外を上書きしない
- 不明:
  - helper 非経由 close 全件の一般整理はこの文書では完了していない。ただし、会計フォーム / import prompt の「Excelを閉じる」直 `workbook.Close()` は `docs/accounting-close-lifecycle-current-state.md` で安定化済み契約として固定済みであり、残課題扱いしない。

### 時系列フロー

1. HOME close では `KernelWorkbookService.RequestCloseHomeSessionFromForm(...)` が backend close を調停し、pending / rejected の間は `FormClosing` を cancel する。
2. backend close 成功後にだけ `FinalizePendingHomeSessionCloseAfterFormClosed()` が走り、HOME session / binding / visibility を解放する。
3. CASE / Base close では `CaseWorkbookBeforeClosePolicy` が `Ignore` / `SuppressPromptForManagedClose` / `PromptForDirtySession` / `SchedulePostCloseFollowUp` を決める。
4. dirty path は `before-close -> dirty prompt -> folder offer -> managed close -> post-close follow-up` の順を守る。
5. managed close 実行中は `ManagedCloseState` スコープで prompt を抑止し、`WorkbookCloseInteropHelper.CloseWithoutSave(workbook)` を使って named argument を避けた close を行う。
6. `PostCloseFollowUpScheduler` は visible workbook 不在時だけ `Quit` を試み、成功後は終了中 `Application` を restore しない。
7. hidden reflection session では `CloseWorkbookQuietly(...)` が `CloseWithoutSave -> FinalRelease(workbook)`、`QuitApplicationQuietly(...)` が `Quit -> FinalRelease(application)` の順を持つ。
8. 会計フォーム / import prompt の「Excelを閉じる」経路では、form / prompt の `Close()` / `Dispose()` と黄色セル cleanup を先に行い、その後に直 `workbook.Close()` を呼ぶ。`Workbooks.Count == 0` の場合だけ `Application.Quit()` へ進む。
9. temporary COM object は owner 側で局所解放する。例として `MoveInitialCursorToHomeCell(...)` の `Range`、`ApplyBaseHomeKeyValues(...)` の `Range`、`AccountingSetKernelSyncService` の owned workbook がある。

### 責務境界

- HOME close fail-closed handshake:
  - `KernelWorkbookCloseService`
- CASE dirty / managed close orchestration:
  - `CaseWorkbookLifecycleService`
- prompt / folder offer:
  - `CaseClosePromptService` / `CaseFolderOpenService`
- post-close quit:
  - `PostCloseFollowUpScheduler`
- hidden session owned cleanup:
  - `KernelUserDataReflectionService`
- temporary COM release:
  - 各 owner service の finally

### 現在の ownership 混在

- close 本線は `KernelWorkbookCloseService`、`CaseWorkbookLifecycleService`、`PostCloseFollowUpScheduler` に分かれる一方、HOME close fail-closed と CASE close ordering が同じ lifecycle の注意事項として残っている。
- COM release timing は 1 箇所に集約されておらず、hidden session owner、temporary Range owner、owned workbook owner に分散している。
- `CaseWorkbookLifecycleService` には close 本線と CASE HOME 表示補正が同居している。

### 危険領域

- fail-closed-sensitive
- COM-sensitive
- ordering-sensitive
- shutdown-sensitive
- busy-retry-sensitive

### 将来の切り出し候補

- close handshake inventory
- owned workbook cleanup helper catalog
- temporary COM scope helper
- post-close quit policy inventory

## 8. Accounting workbook open / sync

### フロー概要

- 入口:
  - `AccountingSetCreateService.Execute(caseWorkbook)`
  - `AccountingSetKernelSyncService.Execute(kernelWorkbook)`
- source-of-truth:
  - `AccountingSetCreateService`
  - `AccountingSetKernelSyncService`
  - `docs/flows.md`
- context:
  - CASE context
  - `SYSTEM_ROOT`
  - accounting template path
  - current/shared `Excel.Application`
- fail-closed 条件:
  - CASE データ、template path、output path、current application が欠ける場合は進めない
  - kernel sync は不要な別 `Excel.Application` fallback を再導入しない
- 不明:
  - 会計シート上の各セル値の業務意味はコードだけでは確定しない

### 時系列フロー

1. CASE create では `AccountingSetCreateService` が CASE context、template path、output folder、output path を解決する。
2. `File.Copy(...)` が最初の実副作用になり、その後 `SuppressPath(...)` を設定して current application で workbook を open する。
3. workbook open 後に `SetWorkbookWindowsVisible(true)` を行い、`BeginInitializationScope()` の内側で DocProperty 設定、初期セル書込、代理人反映をまとめて行う。
4. 初期化 scope は `ScreenUpdating` と `EnableEvents` を退避して `false` にし、dispose で元値へ戻す。
5. 初期化後に suppression を release し、`ActivateInvoiceEntry(...)` を実行し、`ShowWorkbookTaskPaneWhenReady(...)` へ handoff する。
6. create 失敗時は opened workbook を close し、suppression を release し、作成済み output file を delete する。
7. Kernel sync では user-data snapshot から transfer plan を作り、対象 workbook が既に open なら再利用して save し、close はしない。
8. 未 open の場合は `kernelWorkbook.Application` を shared/current app として使い、`ExcelApplicationStateScope` で alerts / screenUpdating / events を抑制し、`OpenInCurrentApplication(...)` で open して hidden window のまま反映 / save / quiet close する。
9. kernel sync の owned workbook cleanup は `CloseWorkbookQuietly(...)` で `CloseWithoutSave -> FinalRelease(workbook)` の順で閉じる。

### 責務境界

- CASE create plan:
  - `AccountingSetCreateService`
- workbook open / visible / cell write / save-as:
  - `AccountingWorkbookService`
- ready-show handoff:
  - `TaskPaneRefreshOrchestrationService`
- kernel-to-accounting transfer plan:
  - `AccountingSetKernelSyncService`
- application state restore:
  - `BeginInitializationScope()` / `ExcelApplicationStateScope`

### 現在の ownership 混在

- `AccountingSetCreateService` が path resolve、wait UI、file copy、副作用 cleanup、initialization、ready-show handoff をまとめて持つ。
- `AccountingSetKernelSyncService` が transfer plan build、shared app quiet scope、open/hidden/save/close をまとめて持つ。
- accounting workbook の cell write owner は `AccountingWorkbookService` だが、いつ visible にするか、いつ hidden のまま save するかは upstream service が決めている。

### 危険領域

- shared-app-sensitive
- ScreenUpdating/EnableEvents-sensitive
- fail-closed-sensitive
- COM-sensitive
- window-sensitive

### 将来の切り出し候補

- create output plan builder
- accounting initialization block service
- kernel sync transfer-plan service
- accounting workbook cleanup helper

## 9. reflection / navigation

### フロー概要

- 入口:
  - `KernelUserDataReflectionService`
  - `DocumentCommandService` の `caselist`
  - HOME からの `ShowKernelSheetAndRefreshPaneFromHome(...)`
- source-of-truth:
  - `KernelUserDataReflectionService`
  - `DocumentCommandService`
  - `ThisAddIn.ShowKernelSheetAndRefreshPaneFromHome(...)`
- context:
  - `WorkbookContext`
  - bound Kernel workbook
  - `SYSTEM_ROOT`
  - source workbook role
- fail-closed 条件:
  - reflection は `WorkbookContext`、Kernel workbook、`SYSTEM_ROOT` が一致しないと進めない
  - navigation は bound context / resolved Kernel workbook が無ければ進めない
- 不明:
  - navigation UX の正式仕様はコードだけでは全件確定しない

### 時系列フロー

1. reflection では `ResolveReflectionKernelWorkbook(context)` が `context == null`、`context.Workbook == null`、非 Kernel、`SYSTEM_ROOT` 不一致を例外扱いし、補正せず止める。
2. shared Excel 側では `ExcelApplicationUiState.Capture(...)` で `ScreenUpdating`、`EnableEvents`、`DisplayAlerts`、`StatusBar` を退避し、quiet mode を適用する。
3. `shUserData` から snapshot を読み取り、Base / Accounting それぞれに open 済み workbook 再利用か hidden reflection session を選ぶ。
4. hidden reflection session は `CreateHiddenIsolatedApplication() -> OpenWorkbookInManagedHiddenSession() -> SetWorkbookWindowsVisible(false) -> Apply plan -> RestoreOwnedWorkbookWindowVisibilityForSave() -> Save() -> CloseWorkbookQuietly() -> QuitApplicationQuietly()` で owner 内 cleanup を完結する。
5. navigation では CASE からの案件一覧登録後、`CreateKernelSheetTransitionContext(...)` が Kernel workbook と `SYSTEM_ROOT` を持つ `WorkbookContext` を作る。
6. `ShowKernelSheetAndRefreshPaneFromHome(...)` は `ResolveKernelWorkbook(context)` に成功した場合だけ `HideKernelHomePlaceholder() -> TryShowSheetByCodeName(...) -> RefreshTaskPane(...)` の順で遷移する。
7. HOME 表示直後の activate 系イベント再入を避けるため、navigation 前に `SuppressUpcomingKernelHomeDisplay(...)` を入れる。

### 責務境界

- reflection precondition / context validation:
  - `KernelUserDataReflectionService`
- Base / Accounting write plan:
  - `KernelUserDataReflectionService`
- hidden reflection session owner:
  - `KernelUserDataReflectionService`
- CASE -> Kernel navigation:
  - `DocumentCommandService` + `ThisAddIn`
- Kernel sheet show:
  - `KernelWorkbookService`

### 現在の ownership 混在

- `KernelUserDataReflectionService` が context validation、quiet mode、hidden session ownership、Base plan、Accounting plan、COM release timingを同時保持している。
- navigation は `DocumentCommandService`、`ThisAddIn`、`KernelWorkbookService` にまたがり、UI transition と workbook resolve が 1 箇所に閉じていない。
- reflection と navigation はどちらも Kernel 文脈を起点にするが、hidden session owner と UI transition owner は別れている。

### 危険領域

- context-sensitive
- fail-closed-sensitive
- COM-sensitive
- ScreenUpdating-sensitive
- window-sensitive

### 将来の切り出し候補

- reflection session owner abstraction
- Base / Accounting reflection plan service
- Kernel navigation coordinator
- shared app quiet-mode helper catalog

## 10. publication / template sync

### フロー概要

- 入口:
  - `KernelCommandService -> KernelTemplateSyncService.Execute(context)`
- source-of-truth:
  - `KernelTemplateSyncService`
  - `PublicationExecutor`
  - `docs/architecture.md`
  - `docs/taskpane-architecture.md`
- context:
  - `WorkbookContext`
  - resolved Kernel workbook
  - `SYSTEM_ROOT`
  - `CaseList_FieldInventory`
- fail-closed 条件:
  - `WorkbookContext` 必須
  - resolved Kernel workbook が無い場合は中止
  - preflight failure では副作用を起こさない
  - kernel save failure では Base sync / invalidate へ進めない
- 不明:
  - publication 全体の transaction / rollback 仕様は現行 docs にはない

### 時系列フロー

1. `KernelTemplateSyncService.Execute(context)` が `context` を必須とし、`ResolveKernelWorkbook(context)` で対象 Kernel workbook を確定する。
2. `ExcelApplicationStateScope` で `ScreenUpdating=false`、`EnableEvents=false` を適用し、master sheet を取得して一時保護解除 scope を張る。
3. `ResolveSystemRoot(...)` と `LoadDefinedTemplateTags(...)` を行い、`KernelTemplateSyncPreflightService.Run(...)` で validation preflight を実行する。
4. preflight が失敗した場合は failure result を返し、副作用を発生させない。
5. success 時だけ `PublicationExecutor.PublishValidatedTemplates(...)` が `WriteToMasterList -> TASKPANE_MASTER_VERSION +1 -> Kernel save -> BuildTaskPaneSnapshot -> Base snapshot sync -> InvalidateCache` の順で side effects を実行する。
6. `SaveKernelWorkbook(...)` が publication commit boundary であり、Base sync はこの後にしか走らない。
7. Base snapshot sync は Base が未 open なら current application で open し、save 後に managed close scope 付きで close する。
8. base sync failure では invalidate は実行し、success + warning を返す current-state を維持する。

### 責務境界

- context-bound Kernel resolve:
  - `KernelWorkbookService`
- preflight:
  - `KernelTemplateSyncPreflightService`
- publication side effects:
  - `PublicationExecutor`
- Base snapshot storage:
  - `KernelTemplateSyncService.SaveSnapshotToBaseWorkbook(...)`
- cache invalidate:
  - `MasterTemplateCatalogService`

### 現在の ownership 混在

- `KernelTemplateSyncService` が UI/application state scope、sheet protection restore、preflight orchestration、result build を持つ。
- `PublicationExecutor` が master list write、version bump、kernel save、Base sync、invalidate を 1 つの side-effect owner として持つ。
- Base sync close path が publication でありながら lifecycle の managed close scope に接続している。

### 危険領域

- SYSTEM_ROOT-sensitive
- fail-closed-sensitive
- ordering-sensitive
- COM-sensitive
- sheet-protection-sensitive

### 将来の切り出し候補

- preflight orchestration inventory
- Base snapshot sync service
- publication result / message builder separation
- publication close-path helper inventory

## この文書で固定すること

- 「今こう動いている」順序と fail-closed 条件
- `WorkbookOpen -> WorkbookActivate -> WindowActivate` 境界
- `WorkbookContext` / `SYSTEM_ROOT` が source-of-truth である経路
- hidden session が例外であること
- ownership 混在を可視化すること

## この文書で固定しないこと

- あるべき設計
- service 新設
- runtime surgery
- retry 数値や protection 秒数の設計意図の断定
- 実機観測なしでの UX 正式仕様化
