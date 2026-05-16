# ThisAddIn Boundary Inventory

## 位置づけ

この文書は、現行 `main` にある `dev/CaseInfoSystem.ExcelAddIn/ThisAddIn.cs` の責務を棚卸しし、今後の安全な境界整理に備えるための inventory です。

- 構成全体の前提: `docs/architecture.md`
- 主要フローの前提: `docs/flows.md`
- UI 制御の前提: `docs/ui-policy.md`
- TaskPane 現行設計の前提: `docs/taskpane-architecture.md`
- TaskPane 現在地の補足: `docs/taskpane-refactor-current-state.md`
- 実機テスト観点: `docs/thisaddin-startup-test-checklist.md`

この文書の目的は、`ThisAddIn` を今すぐ分割することではありません。振る舞い不変を前提に、どの責務が add-in 境界に残っているか、どこが高危険度か、どの単位なら次に小さく切れるかを明確にすることです。

## 1. この文書の目的

- `ThisAddIn` を「巨大クラスだからすぐ分割する」ための文書ではなく、「安全に 1 責務ずつ切り出すための現在地メモ」として残す
- lifecycle / application event / TaskPane / Kernel HOME / COM automation の責務境界を混同しない
- `WorkbookOpen` と window 安定境界を混同しない
- 次回以降の CODEX 作業で、危険領域を避けた最小実装単位を選びやすくする

## 2. 現在の ThisAddIn の責務

### Startup / Shutdown lifecycle

- `ThisAddIn_Startup(...)` が logger 初期化、診断 trace、`AddInCompositionRoot` compose、依存 field の適用を行う
- Startup 周辺は private helper で呼び出しの見通しだけ整理済みだが、`logger 初期化 -> trace -> compose -> 依存適用 -> event 初期化 -> hook -> startup context 判定 / Kernel HOME 表示判定 -> startup refresh` の順序と lifecycle 責務は `ThisAddIn` に残す
- startup 時に Excel application event を購読する
- startup 時に `TryShowKernelHomeFormOnStartup()` と `RefreshTaskPane("Startup", null, null)` を起動する
- `ThisAddIn_Shutdown(...)` が event unhook、pending pane refresh timer 停止、Kernel HOME form close、`TaskPaneManager.DisposeAll()`、word warm-up timer 停止、retained hidden app-cache cleanup (`ShutdownHiddenApplicationCache()`) を行う
- `InternalStartup()` が VSTO `Startup` / `Shutdown` への接続を保持する
- `CreateRibbonExtensibilityObject()` が Ribbon 作成の VSTO 境界を保持する

#### Startup 順序固定メモ

以下は現行 `main` の `dev/CaseInfoSystem.ExcelAddIn/ThisAddIn.cs` と関連 service を基準にした Startup 上位順序です。docs 側はこの実コード順序を正として扱います。

1. 起動診断・trace 初期化
   - `InitializeStartupDiagnostics()`
2. composition root / service 群の構成
   - `CreateStartupCompositionRoot()` -> `Compose()` -> `ApplyCompositionRoot()`
3. application event subscription service 初期化
   - `InitializeApplicationEventSubscriptionService()`
4. application event hook
   - `HookApplicationEvents()`
5. startup context 判定
   - `TryShowKernelHomeFormOnStartup()` の中で `KernelWorkbookService.ShouldShowHomeOnStartup()` が `KernelWorkbookStateService`、`KernelStartupContextInspector`、`KernelWorkbookStartupDisplayPolicy` を使って判定する
6. Kernel HOME 表示判定
   - `TryShowKernelHomeFormOnStartup()` が `shouldShow` を確定し、`true` のときだけ `ShowKernelHomePlaceholder()` を呼ぶ
7. 初回 TaskPane refresh
   - `RefreshTaskPane("Startup", null, null)`

private helper 化や読みやすさ改善を行っても、この上位順序は動かさない前提で扱います。

### Application event wiring / unwiring

- `HookApplicationEvents()` は `ApplicationEventSubscriptionService.Subscribe()` を呼び、次の Excel event を既存順序で購読する
  - `WorkbookOpen`
  - `WorkbookActivate`
  - `WorkbookBeforeSave`
  - `WorkbookBeforeClose`
  - `WindowActivate`
  - `SheetActivate`
  - `SheetSelectionChange`
  - `SheetChange`
  - `AfterCalculate`
- `UnhookApplicationEvents()` は `ApplicationEventSubscriptionService.Unsubscribe()` を呼び、同じ event を解除する
- event handler 本体は引き続き `ThisAddIn` に残し、wiring / unwiring だけを薄い専用 service に分離する
- event の順序と対象集合は lifecycle 挙動に影響するため、単なる配線でも add-in 境界の一部になっている

### WorkbookOpen

- `Application_WorkbookOpen(...)` は Kernel 向け trace 開始判定を行った上で、`WorkbookLifecycleCoordinator.OnWorkbookOpen(...)` に委譲する
- `WorkbookOpen` 自体は workbook-only 境界として扱われ、window 確定はここで保証しない

### WorkbookActivate

- `Application_WorkbookActivate(...)` は `WorkbookLifecycleCoordinator.OnWorkbookActivate(...)` への委譲を担当する
- `ThisAddIn` 自体は handler を薄く保っているが、後段で使う protection predicate を add-in 境界に持っている

### WindowActivate

- `Application_WindowActivate(...)` は trace と active state logging を伴う event 境界として残っている
- handler は `WorkbookEventCoordinator.OnWindowActivate(...)` へ委譲する
- `HandleWindowActivateEvent(...)` で `WindowActivatePaneHandlingService.Handle(...)` へ渡す add-in 内部入口を保持している
- `WorkbookOpen -> WorkbookActivate -> WindowActivate` の順序を前提にしている

### WorkbookBeforeClose

- `Application_WorkbookBeforeClose(...)` は cancelable event 境界を保持する
- 実処理は `WorkbookLifecycleCoordinator.OnWorkbookBeforeClose(...)` に委譲する
- close 後の pane 片付けや managed close 連動の入口であるため、薄い handler でも順序依存がある

### TaskPane 連携

- `RequestTaskPaneDisplayForTargetWindow(...)` が force refresh 準備、`PaneDisplayPolicy` 判定、show/hide/reject の分岐、必要時の refresh 呼出しを行う
- `RefreshTaskPane(...)` が trace 付きの refresh 呼出し境界を持ち、実処理は `TaskPaneRefreshOrchestrationService` に委譲する
- `RefreshActiveTaskPane(...)`、`ScheduleActiveTaskPaneRefresh(...)`、`ScheduleWorkbookTaskPaneRefresh(...)`、`ShowWorkbookTaskPaneWhenReady(...)` が ready-show / delayed refresh の入口を保持する
- `CreateTaskPane(...)` / `RemoveTaskPane(...)` が VSTO `CustomTaskPane` の実生成 / 実破棄境界を保持する
- `HasVisibleCasePaneForWorkbookWindow(...)` が visible pane 判定の bridge を持つ
- `SuppressTaskPaneRefresh(...)` が refresh suppression の入退場管理を持つ

### Kernel / CASE 判定

- `IsKernelWorkbook(...)`、`ShouldShowKernelHomeOnStartup(...)` が Kernel 判定・startup 表示判定の add-in 側窓口を持つ
- `HandleKernelWorkbookBecameAvailable(...)` が Kernel workbook 到達後の UI 反映入口を保持する
- `ShouldAutoShowKernelHomeForEvent(...)`、`HandleExternalWorkbookDetected(...)` が Kernel HOME 自動表示 / 外部 workbook 検知の bridge を持つ
- `SuppressUpcomingKernelHomeDisplay(...)`、`ShouldSuppressKernelHomeDisplay(...)`、`ShouldSuppressCasePaneRefresh(...)` が suppression 判定の窓口を持つ
- `BeginCaseWorkbookActivateProtection(...)`、`ShouldIgnoreWorkbookActivateDuringCaseProtection(...)`、`ShouldIgnoreWindowActivateDuringCaseProtection(...)`、`ShouldIgnoreTaskPaneRefreshDuringCaseProtection(...)` が protection 判定の窓口を持つ

### COM / Excel instance 境界

- `RequestComAddInAutomationService()` が COM automation 公開境界を持つ
- `ShowKernelHomeFromAutomation()`、`ReflectKernelUserDataToAccountingSet()`、`ReflectKernelUserDataToBaseHome()` が外部 automation 入口を持つ
- Ribbon 由来の public method 群が `ResolveRibbonTargetWorkbook()` を通じて対象 workbook を解決する
- `ResolveRibbonTargetWorkbook()` は `ActiveWorkbook` が null の場合に「open workbook が 1 冊だけならそれを使う」fallback を持つ
- `ClearKernelSheetCommandCell(...)` が `Application.EnableEvents` の一時変更を含む
- `ReleaseComObject(...)` が COM final release 境界を持つ
- `ResolveWorkbookPaneWindow(...)` は pane 対象 window 解決 bridge として残っている
- word warm-up timer の schedule / stop / tick も add-in 境界で保持している

### ログ / trace

- startup / shutdown / automation / WindowActivate / TaskPane refresh の trace を出力する
- `EnsureKernelFlickerTraceForWorkbookOpen(...)` が Kernel workbook open 時の trace 開始を担う
- `TraceRuntimeExecutionObservation(...)` が実行環境の診断ログを出す
- workbook / window / active state の descriptor helper を保持する

### 既存サービスへの委譲

`ThisAddIn` 自体は業務判断を極力持たず、主処理を既存 service / coordinator に委譲している。ただし、委譲前後の VSTO 境界と UI 境界はまだ残っている。

- `WorkbookLifecycleCoordinator`
- `KernelWorkbookLifecycleService`
- `WindowActivatePaneHandlingService`
- `TaskPaneRefreshOrchestrationService`
- `TaskPaneManager`
- `KernelWorkbookAvailabilityService`
- `KernelHomeCoordinator`
- `KernelHomeCasePaneSuppressionCoordinator`
- `SheetEventCoordinator`

## 3. 危険度仕分け

### 高

- Startup / Shutdown の順序
  - compose、event hook/unhook、timer 停止、pane dispose、retained hidden app-cache cleanup が連動している
- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` の境界
  - 表示順序、window 確定、pane 再表示、suppression/protection に直結する
- TaskPane 表示 / refresh / render / show の入口
  - `RequestTaskPaneDisplayForTargetWindow(...)`
  - `RefreshTaskPane(...)`
  - `ShowWorkbookTaskPaneWhenReady(...)`
- VSTO `CustomTaskPane` 生成 / 破棄境界
  - `CreateTaskPane(...)`
  - `RemoveTaskPane(...)`
  - `TaskPaneHostRegistry` / `TaskPaneHost` と密結合している
- Kernel HOME 表示 / sheet 遷移と suppression/protection 連携
  - `ShowKernelHomePlaceholder(...)`
  - `ShowKernelSheetAndRefreshPane(...)`
  - `ShowKernelHomePlaceholderWithExternalWorkbookSuppression(...)`
- `RunWithScreenUpdatingSuspended(...)`
  - 表示安定化に関与し、`ScreenUpdating` 復元失敗時の扱いも含む

### 中

- `WorkbookBeforeSave` / `WorkbookBeforeClose` の cancelable event 境界
- Ribbon / COM automation 公開入口
- `ResolveRibbonTargetWorkbook()` の fallback 解決
- word warm-up timer
- suppression / protection predicate の proxy 群

### 低

- trace path helper
- workbook / window descriptor helper
- safe getter / safe formatter
- `LogAutomationFailure(...)` のような補助ログ

## 4. 固定する順序と前提

### Startup / Shutdown

- Startup 順序は固定します。現行 `main` では `InitializeStartupDiagnostics() -> CreateStartupCompositionRoot() / ApplyCompositionRoot() -> InitializeApplicationEventSubscriptionService() -> HookApplicationEvents() -> TryShowKernelHomeFormOnStartup() -> RefreshTaskPane("Startup", null, null)` の順で扱います。
- `TryShowKernelHomeFormOnStartup()` では、Kernel HOME 表示判定より前に `KernelStartupContextInspector` と `KernelWorkbookStartupDisplayPolicy` を通る startup context 判定を済ませます。
- Shutdown 順序も固定します。現行 `main` では `UnhookApplicationEvents() -> StopPendingPaneRefreshTimer() -> KernelHomeForm close / null 化 -> TaskPaneManager.DisposeAll() -> word warm-up timer 停止 -> retained hidden app-cache cleanup` の順で扱います。

### Workbook / Window event

- event の前提順序は `WorkbookOpen -> WorkbookActivate -> WindowActivate` です。
- `WorkbookOpen` は workbook-only 境界であり、window 安定境界として扱いません。
- window 依存処理は `WorkbookActivate` / `WindowActivate` 以降を安全境界として扱います。
- `ActiveWorkbook` と `ActiveWindow` は null になりうる前提を維持します。
- Ribbon / automation 経路では `ResolveRibbonTargetWorkbook()` の fallback を維持し、この null 前提を弱めません。

### TaskPane / Kernel HOME

- TaskPane 本線では `refresh -> render -> show / ready-show` の入口順と、suppression / protection が介在する位置を固定します。
- `WorkbookActivate` / `WindowActivate` の host 再利用前提を安易に崩しません。
- `CreateTaskPane(...)` / `RemoveTaskPane(...)` の VSTO 境界を別責務と混ぜません。
- Kernel HOME の表示、退避、suppression、protection の意味は変更しません。
- `ScreenUpdating` を変更した場合は必ず復元します。

### build と実機確認の扱い

- Compile / build 成功は安全なビルド確認であり、runtime `Addins\` 反映や実機での lifecycle 安定確認とは別です。
- 実機確認なしに lifecycle 変更を安全と断定しません。
- 将来候補は候補のまま扱い、実装予定を確定事項のようには書きません。

## 5. 触ってよい境界

- `2026-05-12` の `KernelHomeFormHost` 導入により、code 側の最小候補だった `KernelHomeForm` のインスタンス管理は `ThisAddIn` から切り出し済みです。
- 対象は form の生成、再利用、close、dispose、null guard、表示時の show / activate / bring-to-front ownership に限定します。
- `TryShowKernelHomeFormOnStartup()`、`ShowKernelHomePlaceholder()`、`HideKernelHomePlaceholder()` の呼び出し位置は変えません。
- suppression / protection 判定、Kernel HOME の表示順、TaskPane refresh 本線には入りません。
- `ThisAddIn` は引き続き Startup / Shutdown の呼び出し位置、Kernel HOME 表示入口、TaskPane / KernelWorkbookService との順序調停を保持します。

## 6. 触ってはいけない境界

- Startup pipeline の service 化
- Shutdown pipeline の整理名目での順序変更
- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` の generic handler 化
- event handler 本体変更
- application event 発火順序の前提変更
- TaskPane refresh / render / show / ready-show / suppression / protection 本線の切り替え
- Kernel HOME suppression / protection の意味変更
- `ResolveRibbonTargetWorkbook()` の fallback 削除
- `WorkbookOpen` を window 安定境界として扱う記述
- `ActiveWorkbook` / `ActiveWindow` の null 前提を弱める変更
- `ThisAddIn` を startup / home / pane の大塊に一気に分割する案

## 7. 将来の最小着手候補

- `KernelHomeForm` のインスタンス管理 host 化は `2026-05-12` に最小差分で実装済みです。
- 今後この周辺を触る場合も、呼び出し位置、suppression 判定、表示順序は変えません。
- `ThisAddIn` 側には lifecycle の順序、startup / shutdown の呼び出し位置、Kernel HOME 表示の入口を残します。
- 実装前に `docs/architecture.md`、`docs/flows.md`、`docs/ui-policy.md`、この inventory、現行 `ThisAddIn.cs` を再照合します。
- 実装前に再度 ChatGPT レビューを通します。
- Compile / build 成功と実機確認成功は別扱いのまま維持します。

## 8. 今やらない分離案

- Startup pipeline を段階ごとに service へ分解する案
- Shutdown cleanup を整理名目で再配列する案
- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` を generic event bridge に寄せる案
- `RequestTaskPaneDisplayForTargetWindow(...)`、`RefreshTaskPane(...)`、`ShowWorkbookTaskPaneWhenReady(...)`、`CreateTaskPane(...)` / `RemoveTaskPane(...)` をまとめて切り替える案
- Kernel HOME 表示、退避、suppression / protection をまとめて別塊へ移す案
- `ResolveRibbonTargetWorkbook()` の fallback を削って active workbook 前提へ寄せる案
- `ThisAddIn` を startup / home / pane の大塊に分ける案
- `WorkbookOpen` を安定境界として扱う前提で設計を組み替える案

## 不明として残す事項

- protection の秒数、retry 間隔、ready-show の正式な仕様根拠
- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` の体感差分に関する最終期待挙動
- 実機でのみ観測できるちらつきや表示出遅れの閾値

これらは既存 docs でもコード上の事実までしか確定していないため、この文書でも断定しません。
## 9. TaskPane VSTO Boundary Current-State (2026-05-06)

- `ThisAddIn.CreateTaskPane(...)` and `ThisAddIn.RemoveTaskPane(...)` remain the concrete VSTO adapter boundary for `CustomTaskPane` create/remove.
- `TaskPaneHost` owns the concrete pane instance lifetime once creation happens. The current remove path is still `TaskPaneHost.Dispose()` -> `ThisAddIn.RemoveTaskPane(...)`.
- `TaskPaneHostFactory` is the current control creation and `ActionInvoked` binding owner for Case / Kernel / Accounting panes. Binding is inline and keyed by `windowKey`.
- `TaskPaneHostRegistry` is the replace/register/remove orchestration owner over the shared host map, but it is not the shared host-map owner itself.
- `TaskPaneManager` still owns `_hostsByWindowKey`, so moving VSTO create/remove and moving state ownership are different future tasks and should not be merged.
- Host metadata timing is outside `ThisAddIn`:
  - `TaskPaneManager.RenderHost(...)` writes `WorkbookFullName`.
  - `TaskPaneHostFlowService.RenderAndShowHostForRefresh(...)` writes `LastRenderSignature`.
  - `TaskPaneDisplayCoordinator.HasVisibleCasePaneForWorkbookWindow(...)` and ready-show early-complete consume that metadata.
- Because visible pane early-complete depends on host metadata plus `host.IsVisible`, VSTO boundary cleanup must not be combined with ready-show / retry / foreground recovery work.
- Code comments now mirror this owner map, but current-state behavior remains unchanged. Event unbinding is still implicit in dispose-driven teardown and is tracked as ambiguity/debt only.

## 9.1 B2 Checkpoint Warning (2026-05-06)

- B2 fixed the owner map around the VSTO boundary, but it did not start VSTO lifecycle surgery.
- At this checkpoint:
  - `ThisAddIn` is still only the concrete create/remove adapter boundary,
  - `TaskPaneHost` is still the lifetime holder,
  - `TaskPaneHostFactory` is still the control creation / binding owner,
  - `TaskPaneHostRegistry` is still the replace/register/remove orchestration owner,
  - `TaskPaneManager` is still the shared host-map owner.
- The following remain intentionally untouched and should not be folded into the next task accidentally:
  - create/remove timing
  - metadata timing
  - visible pane early-complete behavior
  - ready-show / retry
  - foreground recovery
  - event unbinding behavior
- Human-side manual smoke for the B2 checkpoint was reported as OK, so this document now treats the current VSTO boundary behavior as externally validated current-state rather than an inferred design target.
- Any next implementation phase that changes `ThisAddIn`, `TaskPaneHost`, or `TaskPaneHostRegistry` behavior should be treated as runtime surgery and should isolate one boundary at a time.

## 9.2 Create/Remove Timing Pointer (2026-05-06)

- The detailed current-state inventory now lives in `docs/taskpane-manager-responsibility-inventory.md` section `B2.6`.
- Current create chain:
  - `ThisAddIn.RequestTaskPaneDisplayForTargetWindow(...)`
  - refresh pipeline
  - `TaskPaneHostRegistry.GetOrReplaceHost(...)`
  - `TaskPaneHostFactory.CreateHost(...)`
  - `TaskPaneHost` constructor
  - `ThisAddIn.CreateTaskPane(...)`
  - registry registration after host construction returns
- Current remove chain:
  - workbook close cleanup, stale Kernel cleanup, incompatible replacement, display failure fallback, and shutdown cleanup all remove through registry/lifecycle flows
  - concrete pane teardown is still `TaskPaneHost.Dispose()` -> `Hide()` -> `ThisAddIn.RemoveTaskPane(...)`
  - explicit event-unbinding ownership is still not present in current state
- This pointer is inventory only. It does not relax the B2 freeze around create/remove timing, metadata timing, ready-show / retry, visibility/foreground behavior, or event unbinding order.

## 9.3 Event Unbinding Pointer (2026-05-07)

- The detailed current-state inventory now lives in `docs/taskpane-manager-responsibility-inventory.md` section `B2.9`.
- `ThisAddIn.UnhookApplicationEvents()` remains the explicit unbind owner for Excel Application events only. It does not unbind TaskPane control `ActionInvoked`.
- `TaskPaneHostFactory` remains the TaskPane control bind owner for Case / Kernel / Accounting hosts.
- `TaskPaneHost.Dispose()` still does `Hide()` -> `ThisAddIn.RemoveTaskPane(...)` -> `_pane = null` with no repo-local `ActionInvoked -= ...` and no explicit control-dispose call in this boundary.
- `ThisAddIn.RemoveTaskPane(...)` remains a thin VSTO adapter that only calls `CustomTaskPanes.Remove(pane)`, so lower-layer control/event disposal timing stays below this boundary and is not asserted from repository code here.
- Compatible host reuse, display-request show-existing, and ready-show early-complete can keep an already-bound host alive without re-entering create/bind.
- This pointer is inventory only. It does not relax the freeze around create/remove timing, metadata timing, ready-show / retry, visibility/foreground behavior, event unbinding order, or `_hostsByWindowKey` ownership.

## 10. ThisAddIn Startup Boundary Current-State (2026-05-16)

- `ThisAddIn` remains the VSTO `Startup` / `Shutdown` entry owner, composition-root entry owner, and Excel Application event wiring / unwiring owner.
- Startup HOME decision handoff, startup refresh handoff, managed-close startup guard facts, empty-startup quit decision, and the `DisplayAlerts` quit bridge now live in `AddInStartupBoundaryCoordinator`.
- Document-action `ScreenUpdating` execution and TaskPane refresh suppression entry / exit now live in `AddInExecutionBoundaryCoordinator`, injected through the composition root as the existing document-command bridge interfaces.
- Runtime execution diagnostic detail building now lives in `AddInRuntimeExecutionDiagnosticsService`.
- This is not a move of VSTO event ownership, Application event subscription ownership, CustomTaskPane create/remove ownership, hidden Excel ownership, TaskPane lifecycle ownership, workbook close ownership, app quit ownership, or COM release ownership.
- Hidden Excel, TaskPane refresh, close lifecycle, and CASE presentation recovery owners remain in their existing services / coordinators.
