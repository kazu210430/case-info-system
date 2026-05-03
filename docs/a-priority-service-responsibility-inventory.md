# Priority A Service Responsibility Inventory

## 目的

この文書は、優先度 A として扱う次の 2 点について、production code を変更せずに現状整理を行うための棚卸しです。

1. 巨大サービスの責務集中の整理
2. App 層からの `ThisAddIn` / `Globals.ThisAddIn` 直接依存の整理

今回の更新では、現行 `main` に取り込み済みの bridge 化を完了済みとして反映し、未着手・保留の領域と別途調査対象を切り分けます。

## 参照した前提 docs

- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/taskpane-refresh-policy.md`

## 今回の対象と非対象

### 対象

- `dev/CaseInfoSystem.ExcelAddIn/App/KernelWorkbookService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneManager.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/CaseWorkbookLifecycleService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/AccountingSetCreateService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/DocumentCreateService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/KernelCasePresentationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshOrchestrationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneDisplayRetryCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WorkbookTaskPaneDisplayAttemptCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WindowActivatePaneHandlingService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/KernelHomeCasePaneSuppressionCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/UI/TaskPaneHost.cs`
- `dev/CaseInfoSystem.ExcelAddIn/AddInCompositionRoot.cs`
- 補足確認:
  - `dev/CaseInfoSystem.ExcelAddIn/App/DocumentCommandService.cs`
  - `dev/CaseInfoSystem.ExcelAddIn/ThisAddIn.cs`

### 非対象

- production code の挙動変更
- テストコード変更
- suppress 条件変更
- retry / attempt logic 変更
- TaskPaneManager 分割の実施
- bridge 実装着手

## 対象フロー要約

- `KernelWorkbookService`
  - `docs/ui-policy.md` の UI 制御方針に沿って、Kernel HOME 表示準備、Kernel workbook の可視/不可視、Excel main window の表示制御を担う。
- `TaskPaneManager`
  - `docs/flows.md` の TaskPane 更新フローで、host 再利用、role 別描画、CASE pane アクション実行、post-action refresh を担う。
- `CaseWorkbookLifecycleService`
  - `docs/flows.md` の CASE クローズフローで、初回初期化、dirty 判定、managed close、post-close follow-up、CASE HOME 表示補正を担う。
- `KernelCasePresentationService`
  - CASE 表示フローで、非表示オープン後の可視化、一時抑止解除、TaskPane ready-show 予約、初期カーソル位置決定を担う。
- `TaskPaneRefreshOrchestrationService` / `WindowActivatePaneHandlingService`
  - `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` / 明示 refresh を入口に、Pane の再描画、遅延表示、Window 解決、protection 中の抑止判定を担う。

## 1. 巨大サービス責務棚卸し

## 1-1. KernelWorkbookService

### 現在担っている責務

- Kernel workbook 解決
  - `GetOpenKernelWorkbook`
  - `ResolveKernelWorkbook`
  - `GetOrOpenKernelWorkbook`
  - `ResolveKernelWorkbookPathFromAvailableSystemRoot`
- Kernel 設定読取・保存
  - `LoadSettings`
  - `SaveNameRuleA`
  - `SaveNameRuleB`
  - `SelectAndSaveDefaultRoot`
- HOME 表示状態管理
  - `PrepareForHomeDisplay`
  - `PrepareForHomeDisplayFromSheet`
  - `CompleteHomeNavigation`
  - `EnsureHomeDisplayHidden`
  - `ReleaseHomeDisplay`
  - `DismissPreparedHomeDisplayState`
- HOME 表示に伴う Excel / workbook window 制御
  - `ApplyHomeDisplayVisibility`
  - `HideExcelMainWindow`
  - `EnsureExcelApplicationVisible`
  - `ShowExcelMainWindow`
  - `HideKernelWorkbookWindows`
  - `ShowKernelWorkbookWindows`
  - `ConcealKernelWorkbookWindowsForHomeDisplay`
  - `SetKernelWindowVisibleFalse`
- HOME 終了時の lifecycle 調停
  - `CloseHomeSession`
  - `SaveAndCloseKernelWorkbook`
  - `CloseKernelWorkbookWithoutLifecycleCore`
  - `RequestManagedCloseFromHomeExitCore`
  - `QuitApplicationCore`

### 責務が集中している箇所

- HOME 表示制御と Excel main window 制御が同一サービスに集中している。
- workbook 解決、設定 I/O、window 最小化/不可視化、lifecycle close 経路が同居している。
- `CloseHomeSession` は save 有無、CASE 作成フロー中判定、直接 close / lifecycle close、HOME 復帰 / Excel 終了を一括で判断している。

### 分割時に守るべき既存挙動

- CASE 作成直後は Kernel workbook を前景へ戻さない。
- `ScreenUpdating` を変更した場合は必ず復元する。
- visible non-kernel workbook がある場合は既存 workbook layout 保護を優先する。
- `WorkbookOpen` 直後の UI 制御追加に寄らない。

### 将来切り出すなら候補になる単位

- `KernelWorkbookAccessService`
- `KernelHomeDisplayService`
- `KernelWorkbookWindowVisibilityService`
- `KernelHomeSessionCloseService`

## 1-2. TaskPaneManager

### 現在担っている責務

- TaskPane host 管理
  - `RegisterHost`
  - `GetOrReplaceHost`
  - `RemoveHost`
  - `RemoveWorkbookPanes`
  - `DisposeAll`
- TaskPane refresh 主調停
  - `RefreshPane`
  - `TryAcceptRefreshPaneRequest`
  - `ResolveRefreshHost`
  - `TryReuseCaseHostForRefresh`
  - `RenderAndShowHostForRefresh`
- 既存 pane 再利用判定
  - `TryShowExistingPane`
  - `TryShowExistingPaneForDisplayRequest`
  - `ShouldShowWithRenderPaneForDisplayRequest`
  - `HasManagedPaneForWindow`
  - `HasVisibleCasePaneForWorkbookWindow`
- host 表示前準備
  - `PrepareHostsBeforeShow`
  - `HideNonCaseHostsExcept`
  - `HideAll`
  - `HideKernelPanes`
  - `HidePaneForWindow`
- role 別描画
  - `RenderHost`
  - `RenderKernelHost`
  - `RenderAccountingHost`
  - `RenderCaseHost`
- CASE pane action 実行
  - `CaseControl_ActionInvoked`
  - `ExecuteCaseAction`
  - `HandleCasePostActionRefresh`
  - `RefreshCaseHostAfterAction`

### 責務が集中している箇所

- host ライフサイクル管理と action 実行が同じクラスに集中している。
- CASE pane の snapshot 解決、ViewState 構築、表示、アクション後 refresh まで 1 クラスに集約されている。
- `RefreshPane` は precondition、host 解決、reuse、render/show を直列で握っている。

### 分割時に守るべき既存挙動

- Window 単位の host 再利用を維持する。
- CASE pane の表示中 host は毎回 version 比較で再生成しない。
- CASE pane action 後の前景維持方針を維持する。
- `DocumentNamePromptService.TryPrepare` を `doc` 実行前にだけ呼ぶ順序を維持する。
- CASE pane 再描画時に selected tab を保持する。

### 将来切り出すなら候補になる単位

- `TaskPaneHostRegistry`
- `TaskPaneRenderService`
- `CasePaneActionService`
- `TaskPaneDisplayPreparationService`

## 1-3. CaseWorkbookLifecycleService

### 現在担っている責務

- CASE / Base 初回初期化の orchestration
- dirty 判定と session 状態管理
- before-close / managed close / post-close follow-up の orchestration
- created case folder offer pending 状態管理
- CASE HOME 表示補正

### 分離済み補助責務

- `CaseClosePromptService`
  - dirty prompt と created case folder offer prompt
- `CaseFolderOpenService`
  - 保存先フォルダ解決、存在確認、Explorer 起動
- `KernelNameRuleReader`
  - Kernel name rule 参照と package 読取
- `ManagedCloseState`
  - managed close の入れ子状態
- `PostCloseFollowUpScheduler`
  - close 後 follow-up、retry、Excel 終了判定

### 責務が集中している箇所

- `CaseWorkbookLifecycleService` 自体は orchestration hub のままで、before-close、session dirty、created case folder offer pending、CASE HOME 表示補正の順序依存を抱える。
- close 本線と CASE HOME 表示補正が同じサービスに同居している。

### 分割時に守るべき既存挙動

- dirty prompt は `保存しますか？` の Yes / No / Cancel を維持する。
- dirty path の `before-close -> dirty prompt -> folder offer -> managed close -> post-close follow-up` 順序を崩さない。
- managed close 中は before-close prompt を抑止する。
- created CASE folder offer は pending マーク済み workbook だけに出す。
- no visible workbook 時だけ Excel を終了する。

## 2. ThisAddIn / Globals.ThisAddIn 直接依存棚卸し

## 2-1. 既存 bridge パターンの確認

`AddInCompositionRoot` と `DocumentCommandService` では、すでに次の bridge パターンが存在する。

- `ThisAddInAccountingSetReadyShowBridge`
- `ThisAddInDocumentCreateHostBridge`
- `ThisAddInScreenUpdatingExecutionBridge`
- `ThisAddInTaskPaneRefreshSuppressionBridge`
- `ThisAddInActiveTaskPaneRefreshBridge`
- `ThisAddInKernelSheetPaneRefreshBridge`
- `ThisAddInWindowActivatePanePredicateBridge`

このため、App 層から `ThisAddIn` の機能へ寄せる既存方式自体は存在する。

## 2-2. 現行 `main` で完了済みの bridge 化

### `AccountingSetCreateService`

- 追加済み bridge
  - `IAccountingSetReadyShowBridge`
  - `ThisAddInAccountingSetReadyShowBridge`
- 元の直接依存
  - `ShowWorkbookTaskPaneWhenReady(...)`
- 現在の呼び出し先
  - `_accountingSetReadyShowBridge.ShowWorkbookTaskPaneWhenReady(...)`
- 現行コードで確認できること
  - bridge は `ThisAddIn.ShowWorkbookTaskPaneWhenReady(...)` への 1:1 委譲である。
  - `AddInCompositionRoot` で bridge を生成して `AccountingSetCreateService` に注入している。
  - この bridge 化は現行 `main` に取り込み済みである。

### `DocumentCreateService`

- 追加済み bridge
  - `IDocumentCreateHostBridge`
  - `ThisAddInDocumentCreateHostBridge`
- 元の直接依存
  - Excel `Application`
  - `StatusBar`
- 現在の呼び出し先
  - `_documentCreateHostBridge.GetApplication()`
  - `_documentCreateHostBridge.SetStatusBar(...)`
  - `_documentCreateHostBridge.ClearStatusBar()`
- 現行コードで確認できること
  - `Application` / `StatusBar` 依存は bridge 経由になっている。
  - bridge は `ThisAddIn.Application` と `StatusBar` 更新を薄く包む構成である。
  - 文書生成フロー自体は `DocumentCreateService` の既存責務から変わっていない。

### `DocumentCommandService`

- 追加済み bridge
  - `IScreenUpdatingExecutionBridge`
  - `ITaskPaneRefreshSuppressionBridge`
  - `IActiveTaskPaneRefreshBridge`
  - `IKernelSheetPaneRefreshBridge`
  - 各 `ThisAddIn...Bridge`
- 元の直接依存
  - `RunWithScreenUpdatingSuspended(...)`
  - `SuppressTaskPaneRefresh(...)`
  - `RefreshActiveTaskPane(...)`
  - `ShowKernelSheetAndRefreshPane(...)`
- 現在の呼び出し先
  - `_screenUpdatingExecutionBridge`
  - `_taskPaneRefreshSuppressionBridge`
  - `_activeTaskPaneRefreshBridge`
  - `_kernelSheetPaneRefreshBridge`
- 現行コードで確認できること
  - `AddInCompositionRoot` で各 bridge を生成して `DocumentCommandService` に注入している。
  - 各 bridge は `ThisAddIn` 側メソッドへの 1:1 委譲である。
  - `Execute(...)` の実行順序を変更したことはコード上から確認できない。

### `WindowActivatePaneHandlingService`

- 追加済み bridge
  - `IWindowActivatePanePredicateBridge`
  - `ThisAddInWindowActivatePanePredicateBridge`
- 元の直接依存
  - `ShouldIgnoreWindowActivateDuringCaseProtection(...)`
- 現在の呼び出し先
  - `_windowActivatePanePredicateBridge.ShouldIgnoreDuringCaseProtection(...)`
- 現行コードで確認できること
  - WindowActivate protection 判定は bridge 経由に置き換わっている。
  - `Handle(...)` の分岐順は `protection 判定 -> external workbook 検出 -> suppression 判定 -> refresh` のまま維持されている。
  - この bridge 化は現行 `main` に取り込み済みである。

## 2-3. 残っている依存箇所一覧

| ファイル | 箇所 | 何のために触れているか | 既存 bridge へ寄せやすさ | 主な注意点 |
| --- | --- | --- | --- | --- |
| `KernelCasePresentationService.cs` | `SuppressUpcomingCasePaneActivationRefresh` / `ShowWorkbookTaskPaneWhenReady` 呼出 | CASE 表示直後の suppression と ready-show 予約 | 中程度 | suppression と ready-show の順序を壊さないこと |
| `TaskPaneRefreshOrchestrationService.cs` | `Globals.ThisAddIn.ShouldIgnoreTaskPaneRefreshDuringCaseProtection` / `HasVisibleCasePaneForWorkbookWindow` | protection 判定と visible pane 早期完了判定 | 中程度 | retry 系の最上流条件なので位置ずれが危険 |
| `TaskPaneRefreshCoordinator.cs` | `Globals.ThisAddIn.BeginCaseWorkbookActivateProtection` | CASE refresh 成功後の protection 開始 | 低い | protection 3入口との組み合わせを崩さないこと |
| `WorkbookLifecycleCoordinator.cs` | `Globals.ThisAddIn.ShouldIgnoreWorkbookActivateDuringCaseProtection` | WorkbookActivate 再入抑止 | 中程度 | activate protection 判定の入口なので timing ずれが危険 |
| `TaskPaneManager.cs` | `TaskPaneHost` 生成時の `ThisAddIn` 注入 / `RequestTaskPaneDisplayForTargetWindow` | host の VSTO 境界と post-action refresh 再表示経路 | 低い | host 管理と表示調停が密結合 |

## 2-4. AddInCompositionRoot から見える境界

### 確認できたこと

- `AccountingSetCreateService`、`DocumentCreateService`、`DocumentCommandService`、`WindowActivatePaneHandlingService` は bridge 経由の境界を現行 `main` に持つ。
- `TaskPaneManager` には `ThisAddIn` 本体を直接渡している。
- `WindowActivatePaneHandlingService` と `TaskPaneRefreshOrchestrationService` には delegate 群を注入しているが、実処理の一部では依然 `Globals.ThisAddIn` に戻る。
- `TaskPaneDisplayRetryCoordinator` と `WorkbookTaskPaneDisplayAttemptCoordinator` は `AddInTaskPaneCompositionFactory` で生成され、`TaskPaneRefreshOrchestrationService` に注入される。

### 整理上の示唆

- `AccountingSetCreateService`、`DocumentCreateService`、`DocumentCommandService`、`WindowActivatePaneHandlingService` は未着手候補ではなく、完了済みとして扱うべきである。
- `KernelCasePresentationService`、`TaskPaneRefreshOrchestrationService`、`TaskPaneRefreshCoordinator`、`WorkbookLifecycleCoordinator` は protection / ready-show / suppression の危険領域として別途調査対象に残る。
- retry / protection / ready-show の確定 policy と未確定事項の切り分けは `docs/taskpane-refresh-policy.md` を正本として読む。
- `TaskPaneManager` と `TaskPaneHost` は `ThisAddIn` による VSTO `CustomTaskPane` 作成境界を持つため、ここは単純な `Globals.ThisAddIn` 排除より一段重い。

## 3. TaskPane suppression 周辺の追加確認

## 3-1. TaskPaneHost 内部利用の確認

### 確認できた事実

- `TaskPaneHost` は `Globals.ThisAddIn` を使っていない。
- `TaskPaneHost` は constructor で受けた `ThisAddIn` を使い、生成時に `CreateTaskPane(...)`、破棄時に `RemoveTaskPane(...)` を呼ぶ。
- `TaskPaneHost` 自身は描画判断を持たず、`Show()` で `PreferredWidth` と `Visible` を設定し、`Hide()` / `Dispose()` で pane を隠すだけの薄い VSTO ラッパーである。

### host 状態・pane 状態の保持内容

- 固定状態
  - `Window`
  - `WindowKey`
  - `UserControl`
  - `ITaskPaneView`
- VSTO 実体
  - `CustomTaskPane _pane`
- `TaskPaneManager` から更新される関連状態
  - `WorkbookFullName`
  - `LastRenderSignature`
- 可視状態参照
  - `IsVisible` は `_pane.Visible` を安全に読む

### TaskPaneManager との責務境界

- `TaskPaneHost`
  - pane の create / show / hide / dispose
  - window 単位の VSTO 実体保持
- `TaskPaneManager`
  - host の作成タイミング
  - role ごとの control 選択
  - `WorkbookFullName` と `LastRenderSignature` の更新
  - host 再利用、visibility 調停、action 後 refresh

### 将来切り出す場合の注意点

- `TaskPaneHost` の `ThisAddIn` 依存は「表示ロジック」ではなく「VSTO `CustomTaskPane` 生成・破棄境界」である。
- `TaskPaneManager.HasVisibleCasePaneForWorkbookWindow(...)` は host の `WorkbookFullName` と `IsVisible` を見て ready-show の早期完了判定に使うため、host metadata を DTO 扱いして失うと retry 挙動が変わる。
- `TaskPaneManager.GetOrReplaceHost(...)` は role 不一致時に既存 host を dispose して差し替えるため、`TaskPaneHost` を単独で切り出すより `HostRegistry + PaneFactory` 方向で分けるほうが安全。

## 3-2. suppression 条件の確認

### `KernelHomeCasePaneSuppressionCoordinator` が何を抑止しているか

- Kernel HOME 側
  - `SuppressUpcomingKernelHomeDisplay(...)`
  - `ShouldSuppressKernelHomeDisplay(...)`
  - `IsKernelHomeSuppressionActive(...)`
- CASE pane 側の activation refresh 抑止
  - `SuppressUpcomingCasePaneActivationRefresh(...)`
  - `ShouldSuppressCasePaneRefresh(...)`
- CASE workbook foreground 回復中の protection
  - `BeginCaseWorkbookActivateProtection(...)`
  - `ShouldIgnoreWorkbookActivateDuringProtection(...)`
  - `ShouldIgnoreWindowActivateDuringProtection(...)`
  - `ShouldIgnoreTaskPaneRefreshDuringProtection(...)`

### 抑止開始・解除の条件

- CASE pane activation refresh 抑止
  - 開始:
    - `KernelCasePresentationService.ShowCreatedCase(...)` の deferred presentation で、
      - transient suppression release
      - workbook window 可視化保証
      - `SuppressUpcomingCasePaneActivationRefresh(workbookFullName, ...)`
      - `ShowWorkbookTaskPaneWhenReady(...)`
    の順で設定される。
  - 条件:
    - 対象 workbook の `FullName` 一致
    - `WorkbookActivate` 用カウント 1 回
    - `WindowActivate` 用カウント 1 回
    - 有効期限 5 秒
  - 解除:
    - `WorkbookActivate` と `WindowActivate` の両カウント消費後
    - または 5 秒経過時

- CASE foreground protection
  - 開始:
    - `TaskPaneRefreshCoordinator.GuaranteeFinalForegroundAfterRefresh(...)` で、CASE refresh 成功後に `BeginCaseWorkbookActivateProtection(...)` を呼ぶ。
  - 条件:
    - role が `Case`
    - workbook full name 非空
    - window hwnd 非空
    - 有効期限 5 秒
  - 解除:
    - コード上、確認できる解除経路は 5 秒経過による期限切れ経路
    - `ClearCaseWorkbookActivateProtection(...)` は private であり、外部から明示解除される公開経路は確認できない

### CASE pane / HOME pane との関係

- 同一 coordinator が Kernel HOME suppression と CASE pane suppression を両方持つ。
- ただし CASE pane suppression は workbook full name ベース、Kernel HOME suppression は event 名とカウントベースで別管理。
- このため「CASE pane suppression を bridge 化するだけ」のつもりでも、coordinator 分離時に HOME 側の外部 workbook 検出経路を巻き込まないよう注意が必要。

### bridge 化時に壊してはいけない挙動

- `KernelCasePresentationService` 側の
  - release
  - workbook window 可視化
  - activation refresh suppression 設定
  - ready-show 予約
  の順序を壊さない。
- protection 判定は `WorkbookActivate` / `WindowActivate` / `TaskPaneRefresh` の 3 入口で揃って効いているため、1 箇所だけ bridge 化して判定タイミングを変えない。
- `ShouldIgnoreTaskPaneRefreshDuringProtection(...)` は「入力 workbook/window が protected target か」ではなく、「現在の active window が protected target か」を見て refresh を無視する。ここを narrower にすると現行の flicker 抑止が変わる。

## 3-3. retry / attempt coordinator の確認

### `TaskPaneRefreshOrchestrationService` が retry / attempt coordinator を使う目的

- hidden open 直後や foreground 回復直後に workbook window / active window / pane host がまだ揃わない時間差を吸収するため。
- ready-show を 1 回で決め打ちせず、
  - 短い遅延の再試行
  - それでも駄目なら 400ms タイマーによる fallback refresh
  を段階的に行うため。
- 既存 visible CASE pane がすでにある場合は、それを成功として早期完了し、不要な再描画を避けるため。

### 実装上の役割分担

- `TaskPaneDisplayRetryCoordinator`
  - `tryShowOnce(..., 1)` を即時実行
  - 失敗時は attempt 2 以降を予約
  - `maxAttempts` 超過で fallback (`ScheduleWorkbookTaskPaneRefresh`) へ移行
- `WorkbookTaskPaneDisplayAttemptCoordinator`
  - 1 回の attempt を
    - workbook window 解決
    - `TryRefreshTaskPane(...)`
    の組として扱う薄い coordinator
- `TaskPaneRefreshOrchestrationService`
  - ready-show 全体の retry、window 可視化補助、保留タイマー、protection 最上流判定を持つ
- `TaskPaneRefreshCoordinator`
  - suppression count 確認
  - workbook window recovery
  - `WorkbookContext` 解決
  - `TaskPaneManager.RefreshPane(...)`
  - CASE refresh 成功後の Word warm-up 予約
  - 最終 foreground 保証
  を担う

### refresh 抑止・再試行・表示安定化との関係

- `TryRefreshTaskPane(...)` の最上流で `Globals.ThisAddIn.ShouldIgnoreTaskPaneRefreshDuringCaseProtection(...)` を見ており、retry 中でも protection が優先される。
- `ResolveWorkbookPaneWindow(...)` は 2 回まで同期的に window 解決を試し、それでも駄目なら retry coordinator 側へ委譲する。
- `ShowWorkbookTaskPaneWhenReady(...)` の ready-show attempt は、初回だけ workbook window を visible に補助する。
- fallback の `_pendingPaneRefreshTimer` は workbook object を見失っても active CASE context が残っていれば active refresh を継続する。

### 仕様として docs に残すべき内容

- ready-show は「即時 1 回で決め打ち」ではなく段階的 retry であること
- visible CASE pane が既にある場合は refresh 不要として成功扱いにすること
- CASE refresh 成功後に `BeginCaseWorkbookActivateProtection(...)` が入ること
- retry は window / host / foreground 安定化の意図が強く見えること
- ただし各 attempt は `TryRefreshTaskPane(...)` を通るため、通常の pane refresh / snapshot 取得経路から完全に切り離されたものとは断定しないこと

### 未確認のまま残すべき内容

- `80ms` / `400ms` / `3 attempts` の値が業務仕様由来か経験則由来かはコードだけでは確定しない。
- `ShouldIgnoreTaskPaneRefreshDuringProtection(...)` が active window 基準で広めに refresh を止める理由は、実装上は確認できるが、設計意図の正式記述は未確認。
- retry / attempt coordinator の正式な設計意図は、コード断面から推測できる範囲を超えては確定しない。
- protection の明示クリア経路が設計上存在するかどうかは未確認。
- retry coordinator を現在の数値以外へ変えた場合の UX 期待値は docs 未記載。

## 4. 現時点の整理

### bridge 化済みとして扱う対象

- `AccountingSetCreateService`
- `DocumentCreateService`
- `DocumentCommandService`
- `WindowActivatePaneHandlingService`

上記 4 本は、現行 `main` では bridge 実装と `AddInCompositionRoot` の配線を確認できるため、未着手候補ではなく完了済みとして扱う。

### 未着手・保留として残す対象

- `KernelCasePresentationService`
- `TaskPaneRefreshOrchestrationService`
- `TaskPaneRefreshCoordinator`
- `WorkbookLifecycleCoordinator`
- `TaskPaneManager`
- `KernelWorkbookService`

上記は、ready-show / suppression / protection / VSTO 境界の危険領域を含むため、今回の docs 更新では完了扱いへ移さず、未着手・保留として残す。

### 別途調査対象として残す事項

- protection 5 秒失効
- retry `80ms`
- fallback timer `400ms`
- retry `3 attempts`
- visible pane early-complete
- ready-show / suppression の順序

### 所見

- `TaskPaneManager` は `TaskPaneHost` の VSTO `CustomTaskPane` 生成・破棄境界を含むため、単純な bridge 化候補としては扱いにくい。
- `KernelCasePresentationService`、`TaskPaneRefreshOrchestrationService`、`TaskPaneRefreshCoordinator`、`WorkbookLifecycleCoordinator` は protection / ready-show / suppression の相互依存が残っており、実装着手前に別途調査が必要である。

## 5. 変更時に守るべき既存挙動まとめ

- `docs/ui-policy.md`
  - `WorkbookOpen` 直後に直接 UI 表示制御を追加しない
  - `ScreenUpdating` は必ず復元する
  - TaskPane は遅延表示前提を崩さない
- `docs/flows.md`
  - CASE 表示後の ready-show 予約順序を壊さない
  - CASE cache / Base cache / Master rebuild の優先順を変えない
  - open 中 CASE の host 再利用方針を崩さない
  - dirty prompt / managed close / post-close follow-up を崩さない
- `docs/architecture.md`
  - TaskPane snapshot / cache は表示補助であり、保存・生成・実行判断の正本にしない
  - allowlist / review の旧 runtime policy 前提へ戻さない

## 6. 関連テスト有無まとめ

| 対象 | テスト状況 |
| --- | --- |
| `KernelWorkbookService` | 専用 policy / thin orchestration テストあり |
| `TaskPaneManager` | 専用 policy / thin orchestration テストあり |
| `CaseWorkbookLifecycleService` | 専用 policy / thin orchestration テストあり |
| `DocumentCreateService` | 専用テストは未確認。`DocumentCommandServiceTests` などから間接参照あり |
| `KernelCasePresentationService` | 専用テスト未確認 |
| `TaskPaneRefreshOrchestrationService` | 専用テスト未確認 |
| `WindowActivatePaneHandlingService` | 専用テスト未確認 |

## 7. 未確認事項

- `TaskPaneRefreshOrchestrationService` の retry 間隔値と最大試行回数の正式な仕様根拠は未確認。
- `KernelHomeCasePaneSuppressionCoordinator` の 5 秒 suppression duration が UX 要件か暫定値かは未確認。
- CASE 表示時の protection 判定を広めに掛けている理由は、コード上の挙動は確認できるが、設計意図の正式文書は未確認。
