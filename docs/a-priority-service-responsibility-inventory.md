# Priority A Service Responsibility Inventory

## 目的

この文書は、優先度Aとして扱う次の2点について、production code を変更せずに現状整理を行うための棚卸しです。

1. 巨大サービスの責務集中の整理
2. App 層からの `ThisAddIn` / `Globals.ThisAddIn` 直接依存の整理

## 参照した前提 docs

- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`

## 今回の対象と非対象

### 対象

- `dev/CaseInfoSystem.ExcelAddIn/App/KernelWorkbookService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneManager.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/CaseWorkbookLifecycleService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/DocumentCreateService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/KernelCasePresentationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshOrchestrationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WindowActivatePaneHandlingService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/AddInCompositionRoot.cs`
- 補足確認:
  - `dev/CaseInfoSystem.ExcelAddIn/App/DocumentCommandService.cs`
  - `dev/CaseInfoSystem.ExcelAddIn/ThisAddIn.cs`

### 非対象

- production code の挙動変更
- テストコード変更
- 即時のサービス分割
- 即時の `Globals.ThisAddIn` 置換

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
- 補助的な workbook / window 状態観測とロギング
  - `RequiresSave`
  - `CountVisibleWorkbooksSafe`
  - `DescribeVisibleNonKernelWorkbookWindows`
  - `DescribeVisibleOtherWorkbookWindows`
  - `DescribeWorkbookWindows`
  - `FormatActiveExcelState`

### 責務が集中している箇所

- HOME 表示制御と Excel main window 制御が同一サービスに集中している。
- workbook 解決、設定 I/O、window 最小化/不可視化、lifecycle close 経路が同居している。
- `CloseHomeSession` は次を一括で判断している。
  - save 有無
  - CASE 作成フロー中か
  - lifecycle service 経由か直接 close か
  - HOME 復帰するか Excel を終了するか
- `ApplyHomeDisplayVisibility` は次を一括で判断している。
  - 非 Kernel workbook が見えているか
  - 既存 window layout を保護するか
  - Kernel window を最小化するか不可視化するか
  - Excel main window 自体を隠すか

### ただちに分割すると危険な箇所

- `CloseHomeSession`
  - CASE 作成完了直後に Kernel を前景へ戻さない分岐が埋め込まれている。
  - `KernelCaseInteractionState.IsKernelCaseCreationFlowActive` と completion action の組合せを壊すと、`docs/ui-policy.md` の表示制御方針に反する可能性が高い。
- `ApplyHomeDisplayVisibility`
  - visible non-kernel workbook 検出時の最小化経路と、Kernel active 時の不可視化経路が分かれている。
  - `WorkbookOpen` 直接依存ではなく専用サービス経由で抑えているため、ここを分解すると UI 制御の呼出順序を壊しやすい。
- `GetOrOpenKernelWorkbook`
  - open 時に `EnableEvents` を抑止し、window をすぐ不可視化している。
  - docs の「WorkbookOpen 依存の UI 制御を追加しない」に関わるため、開き方の差し替えは危険。

### 将来切り出すなら候補になる単位

- 提案:
  - `KernelWorkbookAccessService`
    - workbook 解決、open、path 解決、settings 読取
  - `KernelHomeDisplayService`
    - HOME 表示準備、HOME release、Excel main window 制御
  - `KernelWorkbookWindowVisibilityService`
    - Kernel window 最小化、不可視化、再表示、foreground 補助
  - `KernelHomeSessionCloseService`
    - `CloseHomeSession` の completion action 判定と実行

### 分割時に守るべき既存挙動

- CASE 作成直後は Kernel workbook を前景へ戻さない。
- `ScreenUpdating` を変更した場合は必ず復元する。
- visible non-kernel workbook がある場合は既存 workbook layout 保護を優先する。
- HOME 表示準備中の Kernel workbook は open 済みでも隠し続ける。
- lifecycle service 利用可否で close 経路が変わる挙動を保持する。

### 関連テストの有無

- あり
  - `dev/CaseInfoSystem.Tests/KernelWorkbookServicePolicyTests.cs`
  - `dev/CaseInfoSystem.Tests/KernelWorkbookServiceThinOrchestrationTests.cs`
- 間接利用あり
  - `dev/CaseInfoSystem.Tests/KernelCaseCreationServiceTests.cs`

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
- CASE pane 構築
  - snapshot 取得
  - parse
  - `CaseTaskPaneViewStateBuilder` による view state 化
  - cache 更新通知
- CASE pane action 実行
  - `CaseControl_ActionInvoked`
  - `ExecuteCaseAction`
  - `HandleCasePostActionRefresh`
  - `RefreshCaseHostAfterAction`
  - `RenderCaseHostAfterAction`
- Kernel / Accounting pane action 実行
  - `KernelControl_ActionInvoked`
  - `AccountingControl_ActionInvoked`

### 責務が集中している箇所

- host ライフサイクル管理と action 実行が同じクラスに集中している。
- CASE pane の snapshot 解決、ViewState 構築、表示、アクション後 refresh まで 1 クラスに集約されている。
- `RefreshPane` は precondition、host 解決、reuse、render/show を直列で握っている。
- `CaseControl_ActionInvoked` は UI 起点でありながら、文書名 prompt 準備、文書作成コマンド実行、post-action refresh 方針まで担っている。

### ただちに分割すると危険な箇所

- `GetOrReplaceHost`
  - `TaskPaneHost` 生成時に `ThisAddIn` を渡しており、host の VSTO 境界を暗黙に握っている。
- `RenderCaseHost`
  - snapshot 取得から `DocumentButtonsControl.Render` までが一体で、`docs/flows.md` の CASE cache / Base cache / Master rebuild の見え方に直結する。
- `HandleCasePostActionRefresh`
  - `document create should keep Word in the foreground`
  - `accounting set should keep the generated workbook in the foreground`
  - `case-list` は defer
  - この分岐を壊すと UI 前景維持方針が崩れる。
- `PrepareHostsBeforeShow`
  - Kernel CASE 作成フロー中の non-case host 抑制と、表示前の整理責務が混ざっており、現在の flicker 抑制に効いている可能性が高い。

### 将来切り出すなら候補になる単位

- 提案:
  - `TaskPaneHostRegistry`
    - host 生成、登録、置換、破棄
  - `TaskPaneRenderService`
    - role 別 render、render signature 判定
  - `CasePaneActionService`
    - `doc` / `accounting` / `caselist` 実行と post-action refresh 方針
  - `TaskPaneDisplayPreparationService`
    - `PrepareHostsBeforeShow` と host visibility 調停

### 分割時に守るべき既存挙動

- Window 単位の host 再利用を維持する。
- CASE pane の表示中 host は毎回 version 比較で再生成しない。
- CASE pane action 後の前景維持方針を維持する。
- `DocumentNamePromptService.TryPrepare` を `doc` 実行前にだけ呼ぶ順序を維持する。
- CASE pane 再描画時に selected tab を保持する。

### 関連テストの有無

- あり
  - `dev/CaseInfoSystem.Tests/TaskPaneManagerOrchestrationPolicyTests.cs`
  - `dev/CaseInfoSystem.Tests/TaskPaneManagerThinOrchestrationTests.cs`

## 1-3. CaseWorkbookLifecycleService

### 現在担っている責務

- CASE / Base 初回初期化
  - `HandleWorkbookOpenedOrActivated`
  - `RegisterKnownCaseWorkbookCore`
  - `SyncNameRulesFromKernelToCaseCore`
- dirty 判定と session 状態管理
  - `_sessionDirtyWorkbookKeys`
  - `HandleSheetChanged`
  - `RemoveWorkbookState`
- before-close prompt / managed close
  - `HandleWorkbookBeforeClose`
  - `ShowClosePromptCore`
  - `BeginManagedCloseScope`
  - `ScheduleManagedSessionClose`
  - `ExecuteManagedSessionClose`
- post-close follow-up
  - `SchedulePostCloseFollowUp`
  - `ExecutePendingPostCloseQueue`
  - `SchedulePendingPostCloseRetry`
  - `QuitExcelIfNoVisibleWorkbook`
- created case folder offer
  - `MarkCreatedCaseFolderOfferPending`
  - `PromptToOpenCreatedCaseFolderIfNeeded`
  - `TryPromptToOpenCreatedCaseFolder`
  - `OpenCreatedCaseFolderCore`
- CASE HOME 表示補正
  - `EnsureCaseHomeLeftColumnVisible`
- Kernel name rule 参照と package 読取
  - `ResolveKernelWorkbookPath`
  - `TryGetKernelNameRules`
  - `TryReadKernelNameRulesFromPackage`

### 責務が集中している箇所

- workbook lifecycle と folder follow-up UI が同じサービスに集中している。
- dirty 判定、managed close、Excel 終了判定、CASE HOME 表示補正、Kernel doc property 同期が同居している。
- `HandleWorkbookBeforeClose` は prompt 表示、cancel 制御、folder offer、managed close / post-close follow-up 予約を一括で担っている。

### ただちに分割すると危険な箇所

- `HandleWorkbookBeforeClose`
  - `ref bool cancel` を扱いながら VSTO close 経路に介入している。
  - managed close 中は prompt suppress、dirty 時は cancel + prompt + 後続予約、そうでなければ VSTO follow-up に委譲、という分岐を壊しやすい。
- `ExecuteManagedSessionClose`
  - `DisplayAlerts` 抑止、保存有無、promptless close、post-close follow-up 予約、`BeginManagedCloseScope` が密結合している。
- `ExecutePendingPostCloseQueue`
  - Excel busy retry と `QuitExcelIfNoVisibleWorkbook` が結びついている。
- `SyncNameRulesFromKernelToCase`
  - open Kernel workbook と package 直読の両方を fallback している。

### 将来切り出すなら候補になる単位

- 提案:
  - `CaseWorkbookDirtyStateService`
    - dirty state と workbook state の保持
  - `CaseWorkbookCloseCoordinator`
    - before-close prompt、managed close、cancel 判定
  - `CaseWorkbookPostCloseFollowUpService`
    - post-close queue、retry、Excel 終了判定、folder offer
  - `CaseWorkbookNameRuleSyncService`
    - Kernel から CASE への name rule 同期
  - `CaseHomeWindowLayoutService`
    - `FreezePanes` / `ScrollColumn` 再適用

### 分割時に守るべき既存挙動

- dirty prompt は `保存しますか？` の Yes / No / Cancel を維持する。
- managed close 中は before-close prompt を抑止する。
- created CASE folder offer は pending マーク済み workbook だけに出す。
- no visible workbook 時だけ Excel を終了する。
- CASE HOME の A列可視維持を壊さない。

### 関連テストの有無

- あり
  - `dev/CaseInfoSystem.Tests/CaseWorkbookLifecycleServicePolicyTests.cs`
  - `dev/CaseInfoSystem.Tests/CaseWorkbookLifecycleServiceThinOrchestrationTests.cs`
- 間接利用あり
  - `dev/CaseInfoSystem.Tests/KernelCaseCreationServiceTests.cs`

## 2. ThisAddIn / Globals.ThisAddIn 直接依存棚卸し

## 2-1. 既存 bridge パターンの確認

`AddInCompositionRoot` と `DocumentCommandService` では、すでに次の bridge パターンが存在する。

- `ThisAddInScreenUpdatingExecutionBridge`
  - `RunWithScreenUpdatingSuspended` への橋渡し
- `ThisAddInTaskPaneRefreshSuppressionBridge`
  - `SuppressTaskPaneRefresh` への橋渡し
- `ThisAddInActiveTaskPaneRefreshBridge`
  - `RefreshActiveTaskPane` への橋渡し

このため、App 層から `ThisAddIn` の機能へ寄せる既存方式自体は存在する。

## 2-2. 依存箇所一覧

| ファイル | 箇所 | 何のために触れているか | 分類 | 既存 bridge へ寄せられそうか | すぐ置換すると危険な理由 | 将来候補 |
| --- | --- | --- | --- | --- | --- | --- |
| `DocumentCreateService.cs` | `ExecuteWordCreate` 冒頭と finally | Excel `Application` の `WindowState`、`ScreenUpdating`、`EnableEvents`、`DisplayAlerts`、`Calculation`、`Visible` を退避・変更・復元 | UI制御 + host bridge | 一部 yes。`Application` 直接取得 bridge、`StatusBar` bridge、`ExcelUiState` bridge に分割可能 | 文書作成中の Excel UI 抑止と復元順序を崩すと、`ScreenUpdating` 復元漏れや Excel 表示不整合につながる | `IExcelApplicationBridge` または `IDocumentCreateExcelUiBridge` |
| `DocumentCreateService.cs` | `SetStatusBar` / `ClearStatusBar` | Excel StatusBar 表示 | UI制御 | yes。小さな bridge に分離しやすい | 文書作成進捗表示は補助だが、例外握りつぶしを含むため置換時に呼出点を増やすと差分が広がる | `IExcelStatusBarBridge` |
| `KernelCasePresentationService.cs` | `SuppressUpcomingCasePaneActivationRefresh` 呼出 | CASE 表示直後の pane activation refresh 抑止 | 状態制御 / TaskPane protection bridge | yes。既存の task pane refresh suppression 系と同種の bridge を追加できる | CASE 表示直後の flicker 抑止に直結しており、ready-show 前後の順序を壊すと UI が崩れやすい | `ICasePaneActivationProtectionBridge` |
| `KernelCasePresentationService.cs` | `ShowWorkbookTaskPaneWhenReady` 呼出 | CASE workbook 可視化後の ready-show 予約 | host bridge / UI制御 | yes。`TaskPaneRefreshOrchestrationService` への bridge 化候補 | `Window.Visible = true`、suppression release、ready-show 予約の順序が現状で密結合 | `IWorkbookTaskPaneReadyShowBridge` |
| `TaskPaneRefreshOrchestrationService.cs` | `ShouldIgnoreTaskPaneRefreshDuringCaseProtection` 呼出 | protection 中 refresh 抑止判定 | 状態参照 | yes。predicate bridge に切出し可能 | refresh attempt の最上流分岐なので、置換時に判定位置がずれると suppression 漏れになる | `ITaskPaneRefreshProtectionBridge` |
| `TaskPaneRefreshOrchestrationService.cs` | `HasVisibleCasePaneForWorkbookWindow` 呼出 | ready-show retry 中に既存 visible pane を検出して早期完了する | 状態参照 + host bridge | 部分的に yes。`TaskPaneManager` への reader bridge 化が候補 | `ThisAddIn` を経由して再び `_taskPaneManager` を見ているため、単純置換すると循環依存の組み替えが必要 | `ICasePaneVisibilityReader` |
| `WindowActivatePaneHandlingService.cs` | `ShouldIgnoreWindowActivateDuringCaseProtection` 呼出 | WindowActivate 中の protection 判定 | 状態参照 | yes。predicate bridge 化が候補 | `WindowActivate` は頻発イベントであり、判定移設時に event timing がずれると refresh 暴発のリスクがある | `IWindowActivateProtectionBridge` |

## 2-3. 補足: `ThisAddIn` 直接注入だが今回中心対象外の箇所

### TaskPaneManager

- `TaskPaneHost` 生成時に `_addIn` を渡している。
  - `GetOrReplaceHost`
- action 後 refresh で `_addIn.RequestTaskPaneDisplayForTargetWindow(...)` を呼ぶ。
  - `RefreshCaseHostAfterAction`

分類:

- `TaskPaneHost` 生成時の注入
  - host bridge
- `RequestTaskPaneDisplayForTargetWindow`
  - host bridge + TaskPane 表示調停

所見:

- `TaskPaneManager` は巨大サービス棚卸し対象として要監視。
- ただし、今回の `Globals.ThisAddIn` 中心棚卸しでは、Document/Presentation/Refresh/WindowActivate より優先度は下げてよい。
- `TaskPaneHost` の内部利用詳細は今回未確認。

## 2-4. AddInCompositionRoot から見える境界

### 確認できたこと

- `DocumentCommandService` には bridge 経由の境界をすでに作っている。
- `TaskPaneManager` には `ThisAddIn` 本体を直接渡している。
- `WindowActivatePaneHandlingService` と `TaskPaneRefreshOrchestrationService` には、`ThisAddIn` ではなく delegate 群を渡しているが、実処理の一部で依然 `Globals.ThisAddIn` へ戻っている箇所がある。

### 整理上の示唆

- `DocumentCreateService` は `DocumentCommandService` と同じ composition 単位に属しているため、bridge 化を足すなら既存パターンに最も寄せやすい。
- `TaskPaneRefreshOrchestrationService` / `WindowActivatePaneHandlingService` は、composition 上は delegate 注入済みなので、`Globals.ThisAddIn` 依存だけを追加 bridge へ寄せる余地がある。
- `KernelCasePresentationService` は現在 root で直接 bridge を受けていないため、ready-show / protection 系の細い bridge を追加する場合は constructor 変更が必要。

## 3. 今後の安全な着手順案

以下は実装提案であり、現時点では推測を含む。

## 3-1. 影響範囲が小さい順

1. `DocumentCreateService` の `StatusBar` / Excel UI state 参照を bridge 化する
2. `WindowActivatePaneHandlingService` の protection 判定を bridge 化する
3. `TaskPaneRefreshOrchestrationService` の protection 判定を bridge 化する
4. `KernelCasePresentationService` の ready-show / suppression 呼出を bridge 化する
5. `TaskPaneManager` の `ThisAddIn` 直接注入用途を host bridge / display request bridge に分ける
6. `CaseWorkbookLifecycleService` の post-close follow-up 単位を分離検討する
7. `KernelWorkbookService` の HOME display / window visibility 単位を分離検討する

## 3-2. 事故リスクが低い順

1. `DocumentCreateService` の `StatusBar` 制御
2. `DocumentCreateService` の `Application` 参照集約
3. `WindowActivatePaneHandlingService` の protection predicate 抽出
4. `TaskPaneRefreshOrchestrationService` の protection predicate 抽出
5. `KernelCasePresentationService` の ready-show bridge 化
6. `TaskPaneManager` の post-action refresh bridge 化
7. 巨大サービスの内部責務分離

## 3-3. 設計改善効果が大きい順

1. `TaskPaneManager` の host 管理 / render / action 実行の分離
2. `KernelWorkbookService` の HOME display / workbook access / window visibility 分離
3. `CaseWorkbookLifecycleService` の close coordinator / post-close follow-up 分離
4. `TaskPaneRefreshOrchestrationService` と `WindowActivatePaneHandlingService` の `Globals.ThisAddIn` 排除
5. `DocumentCreateService` の Excel host bridge 化

## 4. 変更時に守るべき既存挙動まとめ

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

## 5. 関連テスト有無まとめ

| 対象 | テスト状況 |
| --- | --- |
| `KernelWorkbookService` | 専用 policy / thin orchestration テストあり |
| `TaskPaneManager` | 専用 policy / thin orchestration テストあり |
| `CaseWorkbookLifecycleService` | 専用 policy / thin orchestration テストあり |
| `DocumentCreateService` | 専用テストは未確認。`DocumentCommandServiceTests` などから間接参照あり |
| `KernelCasePresentationService` | 専用テスト未確認 |
| `TaskPaneRefreshOrchestrationService` | 専用テスト未確認 |
| `WindowActivatePaneHandlingService` | 専用テスト未確認 |

## 6. 未確認事項

- `TaskPaneHost` が `ThisAddIn` を内部でどう使うかは今回未確認。
- `KernelHomeCasePaneSuppressionCoordinator` の全 suppress 条件は今回未確認。
- `TaskPaneRefreshOrchestrationService` の retry / attempt coordinator の詳細設計意図は docs 未記載であり、コード断面からの把握に留まる。
