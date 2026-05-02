# ThisAddIn Boundary Inventory

## 位置づけ

この文書は、現行 `main` にある `dev/CaseInfoSystem.ExcelAddIn/ThisAddIn.cs` の責務を棚卸しし、今後の安全な境界整理に備えるための inventory です。

- 構成全体の前提: `docs/architecture.md`
- 主要フローの前提: `docs/flows.md`
- UI 制御の前提: `docs/ui-policy.md`
- TaskPane 現行設計の前提: `docs/taskpane-architecture.md`
- TaskPane 現在地の補足: `docs/taskpane-refactor-current-state.md`

この文書の目的は、`ThisAddIn` を今すぐ分割することではありません。振る舞い不変を前提に、どの責務が add-in 境界に残っているか、どこが高危険度か、どの単位なら次に小さく切れるかを明確にすることです。

## 1. この文書の目的

- `ThisAddIn` を「巨大クラスだからすぐ分割する」ための文書ではなく、「安全に 1 責務ずつ切り出すための現在地メモ」として残す
- lifecycle / application event / TaskPane / Kernel HOME / COM automation の責務境界を混同しない
- `WorkbookOpen` と window 安定境界を混同しない
- 次回以降の CODEX 作業で、危険領域を避けた最小実装単位を選びやすくする

## 2. 現在の ThisAddIn の責務

### Startup / Shutdown lifecycle

- `ThisAddIn_Startup(...)` が logger 初期化、診断 trace、`AddInCompositionRoot` compose、依存 field の適用を行う
- Startup 周辺は private helper で呼び出しの見通しだけ整理済みだが、`logger 初期化 -> trace -> compose -> 依存適用 -> event 初期化 -> hook -> Kernel HOME -> startup refresh` の順序と lifecycle 責務は `ThisAddIn` に残す
- startup 時に Excel application event を購読する
- startup 時に `TryShowKernelHomeFormOnStartup()` と `RefreshTaskPane("Startup", null, null)` を起動する
- `ThisAddIn_Shutdown(...)` が event unhook、pending pane refresh timer 停止、Kernel HOME form close、`TaskPaneManager.DisposeAll()`、word warm-up timer 停止、legacy hidden Excel shutdown を行う
- `InternalStartup()` が VSTO `Startup` / `Shutdown` への接続を保持する
- `CreateRibbonExtensibilityObject()` が Ribbon 作成の VSTO 境界を保持する

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
  - compose、event hook/unhook、timer 停止、pane dispose、hidden Excel shutdown が連動している
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

## 4. 絶対に守る境界

- `WorkbookOpen` は window 安定境界ではない
- window 依存処理は `WorkbookActivate` / `WindowActivate` 以降で扱う
- `ActiveWorkbook` は null になりうる前提を維持する
- TaskPane refresh / render / show の順序を壊さない
- `WorkbookActivate` / `WindowActivate` の host 再利用前提を安易に崩さない
- `CreateTaskPane(...)` / `RemoveTaskPane(...)` の VSTO 境界を別責務と混ぜない
- `ScreenUpdating` を変更した場合は必ず復元する
- 実機確認なしに lifecycle 変更を安全と断定しない

## 5. 次に切り出す候補

### 推奨: Application event wiring / unwiring の薄いサービス化

最初に切るなら、`HookApplicationEvents()` / `UnhookApplicationEvents()` だけを薄い registrar へ外出しするのが最も安全です。

現行 branch ではこの候補を `ApplicationEventSubscriptionService` として適用済みです。`ThisAddIn` には startup / shutdown の呼び出し位置と既存 handler 本体を残し、実際の event wiring / unwiring だけを専用 service へ委譲しています。

理由:

- handler 本体を動かさずに、購読 / 解除の責務だけを分離できる
- event の種類と順序を固定したまま整理できる
- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` の処理本線や TaskPane refresh 本線に触れずに済む
- `ThisAddIn` が持つ「VSTO lifecycle」と「Excel application event 配線」を切り分ける第一歩として意味がある

実施時の条件:

- 購読順序を 1 行も変えない
- handler 名も変えない
- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` の処理本体に入らない

### 候補2: trace helper の整理

- `KernelFlickerTrace`、descriptor helper、runtime execution observation を専用 helper へ寄せる
- 低危険度だが、ThisAddIn 境界整理としては効果が限定的

### 候補3: `WorkbookBeforeClose` 周辺の薄い境界整理

- cancelable event 境界を専用 bridge に寄せる候補
- ただし close prompt、managed close、後続 cleanup と密接なため、event wiring より先には勧めない

### 候補4: Startup 初期化順序の読みやすさ改善

- 現行 branch では constructor 引数塊の private helper 抽出まで適用済みで、呼び出し順の見通しだけを改善している
- ただし compose、event hook、startup 表示、startup refresh の順序を動かす分離は引き続き危険であり、次の候補としては扱わない

## 6. 今回は切り出さないもの

### `WorkbookOpen`

- workbook-only 境界として既存前提が強い
- window 安定境界と誤認すると `docs/flows.md` と `docs/ui-policy.md` に反する

### `WorkbookActivate`

- CASE protection、pane refresh suppression、既存 host 再利用と連動している
- `WindowActivate` と分離して単純化しない

### `WindowActivate`

- trace、active state 観測、`WorkbookEventCoordinator`、`WindowActivatePaneHandlingService` の橋渡しがある
- pane 再表示や ready-show に波及しやすい

### TaskPane 表示 / refresh / render / show

- `RequestTaskPaneDisplayForTargetWindow(...)`
- `RefreshTaskPane(...)`
- `ShowWorkbookTaskPaneWhenReady(...)`
- `CreateTaskPane(...)` / `RemoveTaskPane(...)`

これらは現在安定稼働中の TaskPane 系本線に接続しているため、今回切り出さない。

### Excel instance / COM 境界

- `RequestComAddInAutomationService()`
- `ResolveRibbonTargetWorkbook()`
- `ReleaseComObject(...)`
- `ResolveWorkbookPaneWindow(...)`
- hidden Excel shutdown

これらは VSTO / COM / Excel instance の境界そのものであり、切り出しより先に inventory を優先する。

### Kernel HOME 表示 / suppression / protection

- Kernel HOME form 表示と CASE pane suppression は activate 系 event と密接に結び付いている
- docs だけでは最終期待挙動を断定しきれないため、今回切り出さない

## 7. 次回 CODEX 作業への提案

次回に最小実装を投げるなら、次の単位が安全です。

### 提案タスク

`HookApplicationEvents()` / `UnhookApplicationEvents()` だけを担当する薄い service を追加し、`ThisAddIn` から配線責務だけを外へ出す。

### 変更対象の上限

- `dev/CaseInfoSystem.ExcelAddIn/ThisAddIn.cs`
- `dev/CaseInfoSystem.ExcelAddIn/AddInCompositionRoot.cs`
- 新規 service 1 ファイル

### 触ってはいけない範囲

- event handler 本体
- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` の処理順
- TaskPane refresh 本線
- Kernel HOME 表示 / suppression / protection
- `TaskPaneHostRegistry` / `TaskPaneHost` / `TaskPaneManager`

### 完了条件

- 購読 event の種類と順序が完全一致する
- startup / shutdown の呼び出し順が不変である
- コード差分が wiring だけに閉じている
- 実装後の docs 更新も wiring 責務の説明だけに限定する

## 不明として残す事項

- protection の秒数、retry 間隔、ready-show の正式な仕様根拠
- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` の体感差分に関する最終期待挙動
- 実機でのみ観測できるちらつきや表示出遅れの閾値

これらは既存 docs でもコード上の事実までしか確定していないため、この文書でも断定しません。
