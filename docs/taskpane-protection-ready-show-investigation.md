# TaskPane Protection / Ready-Show Investigation

## 目的

この文書は、TaskPane 表示まわりの危険領域について、実装変更前の事実整理を行うための調査メモです。

- policy 正本は `docs/taskpane-refresh-policy.md` とし、この文書は調査メモとして位置づけます。

今回の対象は次の論点です。

- protection 5 秒失効
- retry `80ms` / fallback timer `400ms` / `3 attempts`
- visible pane early-complete
- ready-show / suppression の順序
- `WorkbookActivate` / `WindowActivate` / `TaskPaneRefresh` の protection 判定
- CASE 表示直後の TaskPane 表示順序

## 参照した docs

- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/taskpane-refresh-policy.md`
- `docs/a-priority-service-responsibility-inventory.md`

## 参照したコード

- `dev/CaseInfoSystem.ExcelAddIn/App/KernelCasePresentationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshOrchestrationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WorkbookLifecycleCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WindowActivatePaneHandlingService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneManager.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneDisplayRetryCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WorkbookTaskPaneDisplayAttemptCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/KernelHomeCasePaneSuppressionCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/AddInCompositionRoot.cs`

## 調査対象フローの要約

- `docs/flows.md` では、CASE 表示は `KernelCasePresentationService` を起点とし、Workbook Window 可視化の後に `TaskPaneRefreshOrchestrationService` が TaskPane の準備完了表示を予約する。
- `docs/ui-policy.md` では、TaskPane は遅延表示前提であり、`WorkbookOpen` 直後の直接 UI 表示制御は行わない。
- `docs/architecture.md` では、TaskPane は Window 単位で管理される表示補助であり、snapshot / cache は表示補助として扱う。

## 1. protection の開始地点・終了地点・失効条件

### 確認できた事実

- CASE pane activation suppression は `KernelCasePresentationService.ExecuteDeferredPresentationEnhancements(...)` で、次の順に設定される。
  - `_transientPaneSuppressionService.ReleaseWorkbook(...)`
  - `EnsureWorkbookWindowVisibleBeforeReadyShow(...)`
  - `_casePaneHostBridge.SuppressUpcomingCasePaneActivationRefresh(...)`
  - `_casePaneHostBridge.ShowWorkbookTaskPaneWhenReady(...)`
- この順序は `docs/a-priority-service-responsibility-inventory.md` の既存記述とも一致している。
- suppression の有効期限は `KernelHomeCasePaneSuppressionCoordinator` の `SuppressionDuration = TimeSpan.FromSeconds(5)` を使用する。
- CASE pane activation suppression は対象 workbook の `FullName` を記録し、`WorkbookActivate` 用 1 回、`WindowActivate` 用 1 回のカウントを持つ。
- suppression は次のいずれかで解除される。
  - `WorkbookActivate` と `WindowActivate` の両カウント消費後
  - 5 秒失効時
- CASE foreground protection は `TaskPaneRefreshCoordinator.GuaranteeFinalForegroundAfterRefresh(...)` で、CASE refresh 成功後に `_casePaneHostBridge.BeginCaseWorkbookActivateProtection(...)` を呼ぶことで開始される。
- foreground protection も `SuppressionDuration` を使い、5 秒後に期限切れとなる。
- protection の期限切れ時は `KernelHomeCasePaneSuppressionCoordinator.ClearCaseWorkbookActivateProtection("Expired")` が呼ばれる。

### 補足

- `docs/a-priority-service-responsibility-inventory.md` では、foreground protection の公開解除経路は未確認と整理されている。
- 今回確認したコード断面でも、期限切れ以外の明示解除は見当たらなかった。

## 2. retry / fallback timer の数値と呼び出し元

### 確認できた事実

- `TaskPaneRefreshOrchestrationService` には次の定数がある。
  - `PendingPaneRefreshIntervalMs = 400`
  - `PendingPaneRefreshMaxAttempts = 3`
  - `WorkbookPaneWindowResolveAttempts = 2`
- `WorkbookTaskPaneReadyShowAttemptWorker` には次の定数がある。
  - `ReadyShowRetryDelayMs = 80`
  - `ReadyShowMaxAttempts = 2`
- `KernelCasePresentationService` の `_casePaneHostBridge.ShowWorkbookTaskPaneWhenReady(...)` は、`ThisAddInCasePaneHostBridge` を経由して `TaskPaneRefreshOrchestrationService.ShowWorkbookTaskPaneWhenReady(...)` に到達する。
- `ShowWorkbookTaskPaneWhenReady(...)` は `TaskPaneDisplayRetryCoordinator.ShowWhenReady(...)` を呼び、次を渡す。
  - `TryShowWorkbookTaskPaneOnce`
  - `TaskPaneReadyShowRetryScheduler.Schedule`
  - `StopPendingPaneRefreshTimer`
  - `ScheduleWorkbookTaskPaneRefresh`
- 現行 `main` の `TaskPaneRefreshOrchestrationService` では、`RefreshPreconditionEvaluator`、`RefreshDispatchShell`、`PendingPaneRefreshRetryState`、`WorkbookPaneWindowResolver` への helper split が main に反映済みである。
- `TaskPaneDisplayRetryCoordinator.ShowWhenReady(...)` は次の順に動く。
  - attempt 1 を即時実行
  - 失敗したら attempt 2 以降を `scheduleRetry` で予約
  - `attemptNumber > _maxAttempts` になったら `scheduleFallback` を呼ぶ
- `TaskPaneReadyShowRetryScheduler.Schedule(...)` は ready-show retry delay `80ms` を Timer の `Interval` に設定するため、ready-show retry の遅延は `80ms` である。
- `EnsurePendingPaneRefreshTimer()` は pending timer の `Interval` に `PendingPaneRefreshIntervalMs` を設定しているため、fallback timer の間隔は `400ms` である。
- `ScheduleWorkbookTaskPaneRefresh(...)` は fallback 開始前に一度 `TryRefreshTaskPane(...)` を試し、成功した場合は timer を開始しない。
- `PendingPaneRefreshTimer_Tick(...)` は残回数を 1 ずつ減らしながら retry を続け、成功時または残回数 0 以下で timer を停止する。
- `PendingPaneRefreshTimer_Tick(...)` は対象 workbook を見失った場合でも、active context が CASE なら `TryRefreshTaskPane(_pendingPaneRefreshReason, null, null)` を使って active refresh を継続する。

### 未確認

- `80ms` / `400ms` / `3 attempts` の数値根拠が UX 要件由来か経験則由来かは、コードと既存 docs だけでは確定しない。

## 3. visible pane early-complete の条件

### 確認できた事実

- `TaskPaneRefreshOrchestrationService.TryShowWorkbookTaskPaneOnce(...)` では、window 解決後に `visibleCasePaneAlreadyShown` を判定する。
- 条件は次のとおりである。
  - `resolvedWindow != null`
  - `_casePaneHostBridge.HasVisibleCasePaneForWorkbookWindow(targetWorkbook, resolvedWindow)`
- `TaskPaneManager.HasVisibleCasePaneForWorkbookWindow(...)` 側では、さらに次を確認する。
  - `windowKey` が取得できること
  - その `windowKey` に対応する host が存在すること
  - host の `WorkbookFullName` と対象 workbook の `FullName` が一致すること
  - host の role が `Case` であること
  - host が `Visible` であること
- `visibleCasePaneAlreadyShown` が真の場合、`TryShowWorkbookTaskPaneOnce(...)` は refresh を実行せず、`TaskPaneRefreshAttemptResult.Succeeded()` を返して成功扱いにする。

### 解釈上の注意

- これは「既に表示中の CASE pane が整合しているなら、ready-show retry を refresh 成功相当で終えてよい」という扱いであり、再描画完了そのものとは別である。

## 4. ready-show と suppression の順序

### 確認できた事実

- `KernelCasePresentationService.ExecuteDeferredPresentationEnhancements(...)` の順序は次のとおりである。
  1. workbook に対する transient suppression release
  2. Workbook Window 可視化保証
  3. `SuppressUpcomingCasePaneActivationRefresh(...)`
  4. `ShowWorkbookTaskPaneWhenReady(...)`
  5. 初期カーソル位置調整
- `docs/flows.md` でも、CASE 表示時は Workbook Window 可視化の後に TaskPane の準備完了表示予約へ進む整理になっている。
- `docs/ui-policy.md` の「TaskPane は遅延表示前提」とも整合している。
- `docs/a-priority-service-responsibility-inventory.md` では、この順序を壊してはいけない既存挙動として明記している。

### 未確認

- なぜこの順序でなければならないかの正式な設計意図は、コードと既存 docs だけでは完全には確定しない。

## 5. `WorkbookActivate` / `WindowActivate` / `TaskPaneRefresh` の protection 判定

### 確認できた事実

- `WorkbookLifecycleCoordinator.OnWorkbookActivate(...)` は入口直後に `_casePaneHostBridge.ShouldIgnoreWorkbookActivateDuringCaseProtection(workbook)` を確認し、真なら後続処理へ進まない。
- `WindowActivatePaneHandlingService.Handle(...)` は最初の分岐で `_windowActivatePanePredicateBridge.ShouldIgnoreDuringCaseProtection(workbook, window)` を確認し、真なら後続処理へ進まない。
- `TaskPaneRefreshOrchestrationService.TryRefreshTaskPane(...)` は最上流の `RefreshPreconditionEvaluator` で `_casePaneHostBridge.ShouldIgnoreTaskPaneRefreshDuringCaseProtection(reason, workbook, window)` を確認し、真なら `TaskPaneRefreshAttemptResult.Skipped()` を返す。
- `KernelHomeCasePaneSuppressionCoordinator` 側の挙動は次のとおりである。
  - `ShouldIgnoreWorkbookActivateDuringProtection(...)` は、対象 workbook と active window が protected target と一致する場合に無視する。
  - `ShouldIgnoreWindowActivateDuringProtection(...)` は、対象 workbook と対象 window が protected target と一致する場合に無視する。
  - `ShouldIgnoreTaskPaneRefreshDuringProtection(...)` は protection が active であり、かつ active window が protected target と一致する場合に無視する。

### 注意

- `WorkbookActivate` と `WindowActivate` は入力引数と protected target の一致を見るが、`TaskPaneRefresh` は active window 基準で止める。
- この違いは `docs/a-priority-service-responsibility-inventory.md` にも注意点として残っている。

## 6. CASE 表示直後の TaskPane 表示順序

### 確認できた事実

- `docs/flows.md` では、CASE 表示時の流れは「既知パス登録 -> 一時 suppression -> 非表示オープン等の表示準備 -> Excel ウィンドウ復旧 -> Workbook Window 可視化 -> ready-show 予約 -> 初期カーソル位置移動」で整理されている。
- `KernelCasePresentationService` の実装断面も、少なくとも deferred presentation 部分ではこの順序と整合している。
- `TaskPaneRefreshOrchestrationService.ShowWorkbookTaskPaneWhenReady(...)` は ready-show を即時 1 回で完了させず、`TaskPaneDisplayRetryCoordinator` と `WorkbookTaskPaneDisplayAttemptCoordinator` を通して段階的に待ち合わせる。
- `TryShowWorkbookTaskPaneOnce(...)` では attempt 1 の時だけ `EnsureWorkbookWindowVisibleForTaskPaneDisplay(...)` を呼び、Workbook Window を visible に補助する。
- `TaskPaneRefreshCoordinator` は `KernelFlickerTrace` の structured trace を維持し、`04150a7` で context-resolved / window-resolved / refresh-completed の duplicate plain log が削除済みである。

### 補足

- CASE 表示直後の TaskPane は、`WorkbookOpen` 直後の直接表示ではなく、visible window 解決、retry、fallback を含む遅延表示として扱われている。

## 7. 実機確認が必要な観点

- CASE 作成直後に、Workbook Window 可視化の後で TaskPane が出るか。
- `WorkbookActivate` / `WindowActivate` が連続発火するケースで、protection 5 秒の間に不要な再描画が抑止されるか。
- ready-show retry が 1 回で終わらないケースで、`80ms` retry と `400ms` fallback timer のどちらに入ったか。
- visible CASE pane が既にあるケースで、early-complete により不要な再描画を避けているか。
- target workbook を見失った fallback 中に、active CASE context を使った refresh 継続が実際に起こるか。
- CASE 表示直後のカーソル位置調整と TaskPane 表示が競合しないか。

## 8. まだ実装着手しない方がよい理由

### 確認できた事実

- protection 判定は `WorkbookActivate` / `WindowActivate` / `TaskPaneRefresh` の 3 入口にまたがっている。
- ready-show は `KernelCasePresentationService`、`TaskPaneRefreshOrchestrationService`、`TaskPaneDisplayRetryCoordinator`、`WorkbookTaskPaneDisplayAttemptCoordinator`、`TaskPaneRefreshCoordinator`、`TaskPaneManager` にまたがる。
- `TaskPaneManager` の visible host 状態が early-complete 判定に直接使われている。
- `docs/ui-policy.md` と `docs/flows.md` は、遅延表示・Window 単位再利用・CASE 表示順序の維持を前提にしている。

### ここから言えること

- 1 箇所だけを局所的に変更すると、protection と ready-show の整合を壊す可能性が高い。
- 特に retry 値、fallback timer、suppression 順序、visible pane early-complete は相互に結び付いているため、実装変更前に実機観測を挟むのが妥当である。

## 9. 次に着手するなら最小単位は何か

### 事実ベースで言える範囲

- 実装変更の最小単位をコードだけで断定することはできない。
- 理由は、protection 判定が 3 入口に分散しており、ready-show も複数サービスにまたがっているためである。

### 提案

- 次に着手するなら、まずは「3 入口の protection 判定と CASE 表示直後 ready-show の実機観測手順を固定する」ことが最小単位である。
- 実装変更に入る場合でも、`TaskPaneRefreshOrchestrationService` や retry 値本体ではなく、入口の判定順と実機観測結果の突き合わせから始めるほうが安全に見える。

## 未確認事項

- `80ms` / `400ms` / `3 attempts` の正式な仕様根拠
- protection 5 秒失効が UX 要件か暫定値か
- active window 基準で `TaskPaneRefresh` を止めている正式な設計意図
- fallback timer が実際に必要になる代表ケースの業務上の整理
- CASE 表示直後の各イベント発火順の実機上のばらつき
