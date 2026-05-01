# TaskPaneManager Responsibility Inventory

## 目的

`TaskPaneManager` の現行責務を、`main` に反映済みの実装と既存 docs を前提に棚卸しする。

今回は調査と記録だけを行い、production code は変更しない。

## 参照した docs

- `docs/architecture.md`
- `docs/flows.md`
- `docs/ui-policy.md`
- `docs/taskpane-architecture.md`
- `docs/taskpane-refactor-current-state.md`
- `docs/a-priority-service-responsibility-inventory.md`
- `docs/taskpane-protection-ready-show-investigation.md`

## 調査対象コード

- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneManager.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneDisplayCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneRefreshOrchestrationService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WorkbookLifecycleCoordinator.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/WindowActivatePaneHandlingService.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/TaskPaneBusinessActionLauncher.cs`
- `dev/CaseInfoSystem.ExcelAddIn/App/CasePaneSnapshotRenderService.cs`

## 対象フロー要約

- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` は `WorkbookLifecycleCoordinator` と `WindowActivatePaneHandlingService` を入口にし、TaskPane refresh を直接 `TaskPaneManager` に渡さず、`TaskPaneRefreshOrchestrationService` と `TaskPaneRefreshCoordinator` を経由する。
- window 依存処理は `WorkbookOpen` 直後に確定させず、`ResolveWorkbookPaneWindow(...)` と `EnsurePaneWindowForWorkbook(...)` を通した後で `WorkbookContext` を確定させる。
- `TaskPaneManager` は最終的な host 解決、role 別描画、既存 host 再利用、CASE pane action 後の再描画を担う。
- `doc` 実行前の prompt 初期値準備は `TaskPaneBusinessActionLauncher` が `DocumentNamePromptService.TryPrepare(...)` を先に呼ぶ順序で固定されている。

## 危険度定義

- `A`: 今は触らない方がよい
- `B`: 小さく切り出せる可能性がある
- `C`: docs整理または命名整理で足りる
- `D`: 既に分離済み、または今回の追加対応は不要

## 責務分類

| 分類 | 現在の主担当 | 現状 | 危険度 | 判断 |
| --- | --- | --- | --- | --- |
| Pane生成 | `TaskPaneHostRegistry` | `TaskPaneManager.cs` 内の nested class が `RegisterHost` / `GetOrReplaceHost` / `RemoveHost` / `RemoveWorkbookPanes` / `DisposeAll` を持つ | `B` | VSTO `TaskPaneHost` 生成と action event 配線を 1 箇所に寄せている。責務は閉じているが、まだ `TaskPaneManager.cs` からは分離されていない |
| Pane表示・非表示 | `TaskPaneDisplayCoordinator` | `HideAll` / `HideKernelPanes` / `HideAllExcept` / `HidePaneForWindow` / `TryShowHost` / `PrepareHostsBeforeShow` を担当 | `D` | 主責務は既に別サービスへ分離済み。`TaskPaneManager` 側は薄い委譲のみ |
| Pane再利用 | `TaskPaneDisplayCoordinator`、`TaskPaneRenderStateEvaluator`、`TaskPaneHostReusePolicy`、`TaskPaneManager.TryReuseCaseHostForRefresh(...)` | 既存 host 再表示、render signature 判定、`WorkbookActivate` / `WindowActivate` 時の CASE host 再利用が混在 | `A` | visible pane early-complete と activate 時の再利用が ready-show / protection に直結している |
| Workbook / Window との対応管理 | `WorkbookLifecycleCoordinator`、`WindowActivatePaneHandlingService`、`TaskPaneRefreshOrchestrationService`、`TaskPaneRefreshCoordinator`、`TaskPaneManager.SafeGetWindowKey(...)` | event 入口、window 解決、context 解決、windowKey 単位 host 管理に分散 | `A` | `WorkbookOpen` を window 安定境界にしない前提を壊しやすい |
| TaskPane refresh 起動 | `WorkbookLifecycleCoordinator`、`WindowActivatePaneHandlingService`、`TaskPaneRefreshOrchestrationService`、`TaskPaneRefreshCoordinator` | `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` / ready-show / fallback timer の入口調停を担当 | `D` | refresh 起動責務は `TaskPaneManager` 本体から外へ出ている |
| ViewState / Snapshot 関連 | `ICaseTaskPaneSnapshotReader`、`CasePaneSnapshotRenderService`、`CaseTaskPaneViewStateBuilder`、`TaskPaneSnapshotParser` | snapshot build / parse / view state build は分離済み。`TaskPaneManager` には通知と `Saved` 復元が残る | `B` | 主経路は切れている。残存責務は後処理として小さく切れる |
| Document command / ボタン連携 | `TaskPaneBusinessActionLauncher`、`TaskPaneActionDispatcher`、`TaskPanePostActionRefreshPolicy` | prompt 準備順序は別サービス化済み。post-action refresh 調停は `TaskPaneManager.cs` 内に残る | `B` | button dispatch と post-action refresh は 1 責務として外出ししやすい |
| Excelイベント境界との関係 | `WorkbookLifecycleCoordinator`、`WindowActivatePaneHandlingService`、`TaskPaneRefreshOrchestrationService`、`TaskPaneRefreshCoordinator` | `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` の境界をまたいで動く | `A` | 今回の棚卸し対象ではあるが、次の分割単位にはしない方がよい |
| ログ・診断 | `TaskPaneManager` と周辺 coordinator | `KernelFlickerTrace`、host/context/window descriptor 生成、可視 pane 判定ログ | `C` | まずは docs 上で意味を固定すれば足りる。挙動変更を伴う分離優先度は低い |

## 既に分離済みと見なせるもの

- TaskPane 表示・非表示の主処理は `TaskPaneDisplayCoordinator` に分離済み
- refresh 起動の入口調停は `TaskPaneRefreshOrchestrationService` と `TaskPaneRefreshCoordinator` に分離済み
- WindowActivate 境界の判定は `WindowActivatePaneHandlingService` 側に分離済み
- `doc` 実行前の prompt 初期値準備順序は `TaskPaneBusinessActionLauncher` に分離済み
- CASE pane の snapshot build / parse / view state build は `CasePaneSnapshotRenderService` と関連 reader / builder に分離済み
- CASE cache 更新後処理は `CasePaneCacheRefreshNotificationService` に分離済み
- host reuse / post-action refresh / notification の一部は policy class に切り出されている

## まだ TaskPaneManager に残っているもの

- nested class のまま残っている `TaskPaneHostRegistry`
- nested class のまま残っている `TaskPaneActionDispatcher`
- nested class のまま残っている `TaskPaneRefreshFlowCoordinator`
- `RemoveStaleKernelHosts(...)` による Kernel host の掃除
- `RenderHost(...)` から role 別 render を切り替える最終責務
- host / workbook / window / context の descriptor 生成と trace 出力

## 今は触らない方がよい領域

- `ResolveWorkbookPaneWindow(...)` と `EnsurePaneWindowForWorkbook(...)`
- `HasVisibleCasePaneForWorkbookWindow(...)` を使う visible pane early-complete
- `WorkbookActivate` / `WindowActivate` / `TaskPaneRefresh` をまたぐ protection / suppression
- `WorkbookOpen` を window 安定境界として扱う方向の変更

## 次に切るべき最小単位

### 候補1

`CasePaneCacheRefreshNotificationService` 相当の切り出し

- 対象は `NotifyCasePaneUpdatedIfNeeded(...)`、`TryGetWorkbookSavedState(...)`、`RestoreWorkbookSavedState(...)`
- 1 責務で閉じており、window 境界や ready-show に触れない
- `WorkbookOpen` / `WorkbookActivate` 時だけ通知する現在ルールをそのまま移せる
- 失敗時も戻しやすく、`TaskPaneManager` の render 後処理が見通しよくなる
- 実施済み。render 後にだけ走る副作用境界として、refresh フロー本線や window 解決から分離した

### 候補2

`TaskPaneActionDispatcher` の外出し

- 対象は CASE pane button dispatch と post-action refresh 調停だけ
- `TaskPaneBusinessActionLauncher` と `TaskPanePostActionRefreshPolicy` が既にあるため、責務境界を明示しやすい
- `WorkbookOpen` / `WorkbookActivate` / `WindowActivate` の境界を変えずに進められる

### 候補3

`TaskPaneHostRegistry` の外出し

- 対象は host の生成、差し替え、破棄、workbook 単位の掃除だけ
- VSTO `TaskPaneHost` 生成境界を 1 箇所へ固定しやすい
- ただし `ThisAddIn` と action event 配線を持つため、候補1と候補2よりは慎重に扱う

## 今回の結論

- `TaskPaneManager` は「完全に未分離」ではなく、display / refresh entry / snapshot render / prompt prepare の主責務は既に周辺サービスへ逃がされている
- ただし `TaskPaneManager.cs` には host registry、action dispatch、refresh flow coordinator、CASE cache 更新後処理がまだ残っている
- 次の 1 手は event 境界や ready-show に触らず、render 後処理または action dispatch を小さく外へ出すのが安全

## 不明点

- 既存テストが候補1から候補3をどこまで直接保護しているかは、この調査では網羅確認していない
- visible pane early-complete、retry 値、protection 秒数の正式な仕様根拠は、既存 docs とコードだけでは確定しない
- 実機観測時の体感差分が最も出やすいのが候補2か候補3かは、コードだけでは断定しない

## 設計指針（SOLID 準拠）

本リファクタリングは、振る舞い不変を前提に SOLID 原則へ沿って段階的に進める。

### Single Responsibility Principle（単一責任）

- `TaskPaneManager` は複数責務を持つ状態から、段階的に責務を分離する。
- 「観測」「判定」「副作用」「UI制御」を同じ単位に混在させない。

### Dependency Direction（依存方向）

- 上位のフロー制御は、下位の詳細処理へ直接依存しない形へ寄せる。
- 特に副作用処理は独立サービスとして切り出し、表示調停や event 境界から疎結合に保つ。

### Side Effect Isolation（副作用の隔離）

- 状態変更、キャッシュ更新、通知は明確な副作用境界に閉じる。
- これらを event 境界や window 解決ロジックと混在させない。

### Incremental Refactoring（段階的分離）

- 1回の変更では 1 責務だけを切り出す。
- 常に振る舞い不変を最優先とする。
